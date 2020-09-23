VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_Pay_Employees 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pay Employees"
   ClientHeight    =   6255
   ClientLeft      =   3285
   ClientTop       =   3270
   ClientWidth     =   9735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   9735
   Begin VB.Frame Frame1 
      Height          =   6255
      Left            =   0
      TabIndex        =   60
      Top             =   0
      Width           =   9735
      Begin MSDataGridLib.DataGrid grdDataGrid 
         Bindings        =   "frm_Pay_Employees.frx":0000
         Height          =   3015
         Left            =   120
         TabIndex        =   67
         Top             =   3120
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   5318
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         RowDividerStyle =   5
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
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "PAY"
            Caption         =   "PAY"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "Yes"
               FalseValue      =   "No"
               NullValue       =   "N/A"
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "emp id"
            Caption         =   "EMP ID"
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
            DataField       =   "LASTHOURS"
            Caption         =   "Hours"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "NETCHECK"
            Caption         =   "Check Amount"
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
         BeginProperty Column04 
            DataField       =   "PAYFREQUENCY"
            Caption         =   " Pay Period"
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
         BeginProperty Column05 
            DataField       =   "PAYTYPE"
            Caption         =   "Type"
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
         BeginProperty Column06 
            DataField       =   "LASTDATE"
            Caption         =   "Paid Thru"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   2
            BeginProperty Column00 
               ColumnWidth     =   659.906
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1289.764
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1065.26
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1244.976
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1140.095
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdCreatePayroll 
         Caption         =   "Create Payroll"
         Height          =   375
         Left            =   8160
         TabIndex        =   92
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdPrintChecks 
         Caption         =   "Print Checks"
         Height          =   375
         Left            =   8160
         TabIndex        =   91
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton cmdPyrlItems 
         Caption         =   "Payroll Items"
         Height          =   375
         Left            =   8160
         TabIndex        =   90
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton cmdManualChecks 
         Caption         =   "Manual Check"
         Height          =   375
         Left            =   8160
         TabIndex        =   89
         Top             =   1680
         Width           =   1455
      End
      Begin VB.CommandButton cmdCommissions 
         Caption         =   "Commissions"
         Height          =   375
         Left            =   8160
         TabIndex        =   88
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Frame Frame4 
         Height          =   1215
         Left            =   3000
         TabIndex        =   85
         Top             =   1800
         Width           =   2175
         Begin VB.CheckBox Check1 
            Caption         =   "Manual Check Report"
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   100
            Top             =   960
            Width           =   1935
         End
         Begin VB.OptionButton optCheck 
            Caption         =   "Computer Checks"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   87
            Top             =   240
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.OptionButton optCheck 
            Caption         =   "Manual Checks"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   86
            Top             =   600
            Width           =   1695
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1650
         Left            =   3000
         TabIndex        =   76
         Top             =   120
         Width           =   5055
         Begin VB.TextBox txtFields 
            Height          =   285
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   78
            Top             =   240
            Width           =   2055
         End
         Begin VB.CommandButton cmdLookup 
            Height          =   270
            Left            =   4080
            Picture         =   "frm_Pay_Employees.frx":0010
            Style           =   1  'Graphical
            TabIndex        =   77
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "Bank Account:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   84
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label lblAccounts 
            Height          =   255
            Left            =   600
            TabIndex        =   83
            Top             =   600
            Width           =   3735
         End
         Begin VB.Label Label3 
            Caption         =   "Account Balance:"
            Height          =   255
            Left            =   600
            TabIndex        =   82
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label4 
            Caption         =   "Checks to Print:"
            Height          =   255
            Left            =   720
            TabIndex        =   81
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label lblAcctBal 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   2040
            TabIndex        =   80
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label lblCheckPrints 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   2040
            TabIndex        =   79
            Top             =   1320
            Width           =   1455
         End
      End
      Begin VB.Frame Frame3 
         Height          =   2295
         Left            =   120
         TabIndex        =   68
         Top             =   120
         Width           =   2775
         Begin VB.CommandButton CmdClear 
            Caption         =   "Clear"
            Height          =   375
            Left            =   120
            TabIndex        =   69
            Top             =   1800
            Width           =   2535
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "frm_Pay_Employees.frx":015A
            Left            =   1080
            List            =   "frm_Pay_Employees.frx":016D
            TabIndex        =   72
            Top             =   600
            Width           =   1575
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            ItemData        =   "frm_Pay_Employees.frx":019E
            Left            =   1080
            List            =   "frm_Pay_Employees.frx":01B4
            TabIndex        =   71
            Top             =   960
            Width           =   1575
         End
         Begin VB.CommandButton cmdSelect 
            Caption         =   "Select"
            Height          =   375
            Left            =   120
            TabIndex        =   70
            Top             =   1320
            Width           =   2535
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Pay Select"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   75
            Top             =   240
            Width           =   2535
         End
         Begin VB.Label Label8 
            Caption         =   "Pay Peiod:"
            Height          =   255
            Left            =   120
            TabIndex        =   74
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label9 
            Caption         =   "Type:"
            Height          =   255
            Left            =   120
            TabIndex        =   73
            Top             =   960
            Width           =   975
         End
      End
      Begin VB.TextBox txtFieldsDate 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   66
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox txtFieldsDate 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   65
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox txtFieldsDate 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   64
         Top             =   2640
         Width           =   1095
      End
      Begin VB.CommandButton cmdDate 
         Height          =   285
         Index           =   2
         Left            =   7680
         Picture         =   "frm_Pay_Employees.frx":01F9
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   2640
         Width           =   375
      End
      Begin VB.CommandButton cmdDate 
         Height          =   285
         Index           =   0
         Left            =   7680
         Picture         =   "frm_Pay_Employees.frx":07D3
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   1920
         Width           =   375
      End
      Begin VB.CommandButton cmdDate 
         Height          =   285
         Index           =   1
         Left            =   7680
         Picture         =   "frm_Pay_Employees.frx":0DAD
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox txtFieldsTemp 
         DataField       =   " "
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   93
         Text            =   "231100"
         Top             =   5760
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox TempEMPID 
         Height          =   285
         Left            =   6600
         TabIndex        =   180
         Top             =   2640
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label lblAcctTemp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   840
         TabIndex        =   99
         Top             =   5520
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label ChkDates 
         Caption         =   "Check Date:"
         Height          =   255
         Index           =   0
         Left            =   5280
         TabIndex        =   98
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label ChkDates 
         Caption         =   "Pay Period Starts:"
         Height          =   255
         Index           =   1
         Left            =   5280
         TabIndex        =   97
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label ChkDates 
         Caption         =   "Pay Period Ends:"
         Height          =   255
         Index           =   2
         Left            =   5280
         TabIndex        =   96
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label lblnames 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Employee Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   95
         Top             =   2520
         Width           =   2775
      End
      Begin VB.Label lblName 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   120
         TabIndex        =   94
         Top             =   2760
         Width           =   2775
      End
   End
   Begin VB.Frame frCommision 
      Caption         =   "Commission"
      Height          =   6255
      Left            =   0
      TabIndex        =   41
      Top             =   0
      Visible         =   0   'False
      Width           =   11895
      Begin VB.CommandButton CmdClearComm 
         Caption         =   "Clear All"
         Height          =   375
         Left            =   10560
         TabIndex        =   54
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton cmdSSelectAll 
         Caption         =   "Select All"
         Height          =   375
         Left            =   10560
         TabIndex        =   53
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdEmpLookup 
         Height          =   270
         Left            =   3480
         Picture         =   "frm_Pay_Employees.frx":1387
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   960
         Width           =   375
      End
      Begin VB.CommandButton cmdDateComm 
         Height          =   285
         Left            =   6360
         Picture         =   "frm_Pay_Employees.frx":14D1
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox txtCommission 
         Height          =   285
         Index           =   2
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox txtCommission 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   47
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txtCommission 
         Height          =   285
         Index           =   0
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   960
         Width           =   1815
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Back"
         Height          =   375
         Left            =   10560
         TabIndex        =   43
         Top             =   240
         Width           =   1215
      End
      Begin MSDataGridLib.DataGrid dtgrdDataGrid 
         Bindings        =   "frm_Pay_Employees.frx":1AAB
         Height          =   4455
         Left            =   60
         TabIndex        =   42
         Top             =   1680
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   7858
         _Version        =   393216
         AllowUpdate     =   0   'False
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
         ColumnCount     =   12
         BeginProperty Column00 
            DataField       =   "AR SALE Select to Pay"
            Caption         =   "Pay"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "Yes"
               FalseValue      =   "No"
               NullValue       =   "N/A"
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "EMP Name"
            Caption         =   "Employee Name"
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
            DataField       =   "AR SALE Ext Document #"
            Caption         =   "Document #"
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
         BeginProperty Column03 
            DataField       =   "AR SALE Document Type"
            Caption         =   "Type"
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
         BeginProperty Column04 
            DataField       =   "AR SALE Date"
            Caption         =   "Inv Date"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "COMMPERCENT"
            Caption         =   "%"
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
         BeginProperty Column06 
            DataField       =   "Method"
            Caption         =   "Basis"
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
         BeginProperty Column07 
            DataField       =   "Commission"
            Caption         =   "Commission"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """$""#,##0.00;(""$""#,##0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "AR SALE Customer ID"
            Caption         =   "Customer ID"
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
         BeginProperty Column09 
            DataField       =   "Price"
            Caption         =   "Price"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """$""#,##0.00;(""$""#,##0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column10 
            DataField       =   "Cost"
            Caption         =   "Cost"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """$""#,##0.00;(""$""#,##0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column11 
            DataField       =   "Profit"
            Caption         =   "Profit"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """$""#,##0.00;(""$""#,##0.00)"
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
               ColumnWidth     =   645.165
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1544.882
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1049.953
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   840.189
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   945.071
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   540.284
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   734.74
            EndProperty
            BeginProperty Column07 
               Alignment       =   1
               ColumnWidth     =   945.071
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1065.26
            EndProperty
            BeginProperty Column09 
               Alignment       =   1
               ColumnWidth     =   945.071
            EndProperty
            BeginProperty Column10 
               Alignment       =   1
               ColumnWidth     =   945.071
            EndProperty
            BeginProperty Column11 
               Alignment       =   1
               ColumnWidth     =   945.071
            EndProperty
         EndProperty
      End
      Begin VB.Label lblCommission 
         Caption         =   "Pay Period Ends:"
         Height          =   255
         Index           =   2
         Left            =   7080
         TabIndex        =   49
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lblCommission 
         Caption         =   " CutOff Date:"
         Height          =   255
         Index           =   1
         Left            =   3960
         TabIndex        =   48
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lblCommission 
         Caption         =   "Employee ID:"
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   46
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Select Commisssions to Pay by Invoice"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2400
         TabIndex        =   44
         Top             =   120
         Width           =   6825
      End
   End
   Begin VB.Frame frCheckDetail 
      Height          =   7335
      Left            =   0
      TabIndex        =   101
      Top             =   0
      Visible         =   0   'False
      Width           =   9735
      Begin VB.Frame Frame7 
         Height          =   1215
         Left            =   120
         TabIndex        =   173
         Top             =   120
         Width           =   6255
         Begin VB.CommandButton cmdCheckEmplookup 
            Height          =   270
            Left            =   2640
            Picture         =   "frm_Pay_Employees.frx":1ABB
            Style           =   1  'Graphical
            TabIndex        =   177
            Top             =   360
            Width           =   375
         End
         Begin VB.TextBox txtCheckDetail 
            DataField       =   "EMP ID"
            Height          =   285
            Index           =   0
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   176
            Text            =   " "
            Top             =   360
            Width           =   1695
         End
         Begin VB.CheckBox ckCheckDetail 
            Alignment       =   1  'Right Justify
            Caption         =   "CheckDetail"
            Enabled         =   0   'False
            Height          =   195
            Left            =   3240
            TabIndex        =   175
            Top             =   360
            Width           =   1215
         End
         Begin VB.CommandButton cmdBackCheck 
            Caption         =   "Back"
            Height          =   975
            Left            =   5040
            Picture         =   "frm_Pay_Employees.frx":1C05
            Style           =   1  'Graphical
            TabIndex        =   174
            Top             =   170
            Width           =   1095
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "EMP ID:"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   178
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Frame frCheckDetailsub2 
         Height          =   3015
         Left            =   6480
         TabIndex        =   103
         Top             =   120
         Width           =   3135
         Begin VB.TextBox txtCheckDetail 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            BorderStyle     =   0  'None
            DataField       =   "NETCHECK"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$""#,##0.00;(""$""#,##0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   2
            EndProperty
            Height          =   285
            Index           =   17
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   131
            Text            =   "$"
            Top             =   2640
            Width           =   1695
         End
         Begin VB.TextBox txtCheckDetail 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            BorderStyle     =   0  'None
            DataField       =   "DEDUCTIONS"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$""#,##0.00;(""$""#,##0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   2
            EndProperty
            Height          =   285
            Index           =   16
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   130
            Text            =   "$"
            Top             =   2400
            Width           =   1695
         End
         Begin VB.TextBox txtCheckDetail 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            BorderStyle     =   0  'None
            DataField       =   "ADDITIONS"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$""#,##0.00;(""$""#,##0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   2
            EndProperty
            Height          =   285
            Index           =   15
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   129
            Text            =   "$"
            Top             =   2160
            Width           =   1695
         End
         Begin VB.TextBox txtCheckDetail 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            BorderStyle     =   0  'None
            DataField       =   "LASTLOCAL"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$""#,##0.00;(""$""#,##0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   2
            EndProperty
            Height          =   285
            Index           =   14
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   128
            Text            =   "$"
            Top             =   1920
            Width           =   1695
         End
         Begin VB.TextBox txtCheckDetail 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            BorderStyle     =   0  'None
            DataField       =   "LASTSTATETAX"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$""#,##0.00;(""$""#,##0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   2
            EndProperty
            Height          =   285
            Index           =   13
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   127
            Text            =   "$"
            Top             =   1680
            Width           =   1695
         End
         Begin VB.TextBox txtCheckDetail 
            DataField       =   "LASTFIT"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$""#,##0.00;(""$""#,##0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   2
            EndProperty
            Height          =   285
            Index           =   12
            Left            =   1200
            TabIndex        =   126
            Text            =   "$"
            Top             =   1320
            Width           =   1695
         End
         Begin VB.TextBox txtCheckDetail 
            DataField       =   "LASTFICA"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$""#,##0.00;(""$""#,##0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   2
            EndProperty
            Height          =   285
            Index           =   11
            Left            =   1200
            TabIndex        =   125
            Text            =   "$"
            Top             =   1030
            Width           =   1695
         End
         Begin VB.TextBox txtCheckDetail 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            BorderStyle     =   0  'None
            DataField       =   "LASTAGI"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$""#,##0.00;(""$""#,##0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   2
            EndProperty
            Height          =   285
            Index           =   10
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   124
            Text            =   "$"
            Top             =   720
            Width           =   1695
         End
         Begin VB.TextBox txtCheckDetail 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            BorderStyle     =   0  'None
            DataField       =   "PRETAXDED"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$""#,##0.00;(""$""#,##0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   2
            EndProperty
            Height          =   285
            Index           =   9
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   123
            Text            =   "$"
            Top             =   480
            Width           =   1695
         End
         Begin VB.TextBox txtCheckDetail 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            BorderStyle     =   0  'None
            DataField       =   "LASTGROSS"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$""#,##0.00;(""$""#,##0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   2
            EndProperty
            Height          =   285
            Index           =   8
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   122
            Text            =   "$"
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Net:"
            Height          =   255
            Index           =   11
            Left            =   480
            TabIndex        =   121
            Top             =   2640
            Width           =   615
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Deductions:"
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   120
            Top             =   2400
            Width           =   975
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Additions:"
            Height          =   255
            Index           =   9
            Left            =   240
            TabIndex        =   119
            Top             =   2160
            Width           =   855
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Local Tax:"
            Height          =   255
            Index           =   8
            Left            =   240
            TabIndex        =   118
            Top             =   1920
            Width           =   855
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "State Tax:"
            Height          =   255
            Index           =   7
            Left            =   240
            TabIndex        =   117
            Top             =   1680
            Width           =   855
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "FIT:"
            Height          =   255
            Index           =   6
            Left            =   480
            TabIndex        =   116
            Top             =   1350
            Width           =   615
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "FICA:"
            Height          =   255
            Index           =   5
            Left            =   480
            TabIndex        =   115
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Taxable AGI :"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   114
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Pretax Ded:"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   113
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Gross:"
            Height          =   255
            Index           =   2
            Left            =   480
            TabIndex        =   112
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame frCheckDetailsub1 
         Height          =   1815
         Left            =   120
         TabIndex        =   102
         Top             =   1320
         Width           =   6255
         Begin VB.CommandButton cmdCreatePayrollChk 
            Caption         =   "Create Payroll"
            Height          =   975
            Left            =   5040
            Picture         =   "frm_Pay_Employees.frx":1F0F
            Style           =   1  'Graphical
            TabIndex        =   179
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txtCheckDetail 
            DataField       =   "NETCHECK"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$""#,##0.00;(""$""#,##0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   2
            EndProperty
            Height          =   285
            Index           =   7
            Left            =   4440
            Locked          =   -1  'True
            TabIndex        =   111
            Text            =   " "
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox txtCheckDetail 
            DataField       =   "EMP Postal"
            Height          =   285
            Index           =   6
            Left            =   3240
            Locked          =   -1  'True
            TabIndex        =   110
            Text            =   " "
            Top             =   1320
            Width           =   1095
         End
         Begin VB.TextBox txtCheckDetail 
            DataField       =   "EMP State"
            Height          =   285
            Index           =   5
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   109
            Text            =   " "
            Top             =   1320
            Width           =   1095
         End
         Begin VB.TextBox txtCheckDetail 
            DataField       =   "EMP City"
            Height          =   285
            Index           =   4
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   108
            Text            =   " "
            Top             =   1320
            Width           =   1095
         End
         Begin VB.TextBox txtCheckDetail 
            DataField       =   "EMP Address 2"
            Height          =   285
            Index           =   3
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   107
            Text            =   " "
            Top             =   960
            Width           =   3495
         End
         Begin VB.TextBox txtCheckDetail 
            DataField       =   "EMP Address 1"
            Height          =   285
            Index           =   2
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   106
            Text            =   " "
            Top             =   600
            Width           =   3495
         End
         Begin VB.TextBox txtCheckDetail 
            DataField       =   "Emp Name"
            Height          =   285
            Index           =   1
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   105
            Text            =   " "
            Top             =   240
            Width           =   3495
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Pay To:"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   104
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Frame5 
         Height          =   1815
         Left            =   120
         TabIndex        =   132
         Top             =   3120
         Width           =   6975
         Begin VB.TextBox txtCheckDetail 
            DataField       =   "SS#"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   """$""#,##0.00;(""$""#,##0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            Height          =   285
            Index           =   24
            Left            =   3360
            Locked          =   -1  'True
            TabIndex        =   151
            Text            =   " "
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox txtCheckDetail 
            DataField       =   "LASTHOURS"
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
            Index           =   18
            Left            =   1080
            TabIndex        =   139
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox txtCheckDetail 
            DataField       =   "HOURLYRATE"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$""#,##0.00;(""$""#,##0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   2
            EndProperty
            Height          =   285
            Index           =   19
            Left            =   1080
            TabIndex        =   140
            Text            =   "$"
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox txtCheckDetail 
            DataField       =   "FEDALLOW"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   """$""#,##0.00;(""$""#,##0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            Height          =   285
            Index           =   29
            Left            =   5640
            Locked          =   -1  'True
            TabIndex        =   156
            Text            =   " "
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox txtCheckDetail 
            DataField       =   "AMOUNT"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$""#,##0.00;(""$""#,##0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   2
            EndProperty
            Height          =   285
            Index           =   23
            Left            =   1080
            TabIndex        =   144
            Text            =   "$"
            Top             =   1440
            Width           =   1215
         End
         Begin VB.TextBox txtCheckDetail 
            DataField       =   "LASTCOMMISSION"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$""#,##0.00;(""$""#,##0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   2
            EndProperty
            Height          =   285
            Index           =   22
            Left            =   5640
            Locked          =   -1  'True
            TabIndex        =   143
            Text            =   "$"
            Top             =   1320
            Width           =   1215
         End
         Begin VB.TextBox txtCheckDetail 
            DataField       =   "OTRATE"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$""#,##0.00;(""$""#,##0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   2
            EndProperty
            Height          =   285
            Index           =   20
            Left            =   1080
            TabIndex        =   141
            Text            =   "$"
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox txtCheckDetail 
            DataField       =   "BlankCheckHourlyRate"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$""#,##0.00;(""$""#,##0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   2
            EndProperty
            Height          =   285
            Index           =   31
            Left            =   5640
            TabIndex        =   160
            Text            =   "$"
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox txtCheckDetail 
            DataField       =   "BlankCheckOTRate"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$""#,##0.00;(""$""#,##0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   2
            EndProperty
            Height          =   285
            Index           =   30
            Left            =   5640
            TabIndex        =   158
            Text            =   "$"
            Top             =   840
            Width           =   1215
         End
         Begin VB.TextBox txtCheckDetail 
            DataField       =   "SALARY"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$""#,##0.00;(""$""#,##0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   2
            EndProperty
            Height          =   285
            Index           =   21
            Left            =   5640
            Locked          =   -1  'True
            TabIndex        =   142
            Text            =   "$"
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox txtCheckDetail 
            DataField       =   "FEDFILINGSTATUS"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   """$""#,##0.00;(""$""#,##0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            Height          =   285
            Index           =   25
            Left            =   3360
            Locked          =   -1  'True
            TabIndex        =   152
            Text            =   " "
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox txtCheckDetail 
            DataField       =   "DEPARTMENT"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   """$""#,##0.00;(""$""#,##0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            Height          =   285
            Index           =   26
            Left            =   3360
            Locked          =   -1  'True
            TabIndex        =   153
            Text            =   " "
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox txtCheckDetail 
            DataField       =   "LOCATION"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   """$""#,##0.00;(""$""#,##0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            Height          =   285
            Index           =   27
            Left            =   3360
            Locked          =   -1  'True
            TabIndex        =   154
            Text            =   " "
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox txtCheckDetail 
            DataField       =   "PAYTYPE"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   """$""#,##0.00;(""$""#,##0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            Height          =   285
            Index           =   28
            Left            =   3360
            Locked          =   -1  'True
            TabIndex        =   155
            Text            =   " "
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Hourly Rate:"
            Height          =   255
            Index           =   25
            Left            =   4560
            TabIndex        =   159
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "OT Rate:"
            Height          =   255
            Index           =   24
            Left            =   4560
            TabIndex        =   157
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Fed Allow:"
            Height          =   255
            Index           =   23
            Left            =   4560
            TabIndex        =   150
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Pay Type:"
            Height          =   255
            Index           =   22
            Left            =   2280
            TabIndex        =   149
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Location:"
            Height          =   255
            Index           =   21
            Left            =   2280
            TabIndex        =   148
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Department:"
            Height          =   255
            Index           =   20
            Left            =   2280
            TabIndex        =   147
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Filing Status:"
            Height          =   255
            Index           =   19
            Left            =   2280
            TabIndex        =   146
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Soc Sec #:"
            Height          =   255
            Index           =   18
            Left            =   2280
            TabIndex        =   145
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Amount:"
            Height          =   255
            Index           =   17
            Left            =   240
            TabIndex        =   138
            Top             =   1440
            Width           =   735
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Commission:"
            Height          =   255
            Index           =   16
            Left            =   4560
            TabIndex        =   137
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Salary:"
            Height          =   255
            Index           =   15
            Left            =   4560
            TabIndex        =   136
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "OT Rate:"
            Height          =   255
            Index           =   14
            Left            =   240
            TabIndex        =   135
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Hourly Rate:"
            Height          =   255
            Index           =   13
            Left            =   0
            TabIndex        =   134
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Hours:"
            Height          =   255
            Index           =   12
            Left            =   480
            TabIndex        =   133
            Top             =   240
            Width           =   495
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frm_Pay_Employees.frx":2219
         Height          =   2175
         Left            =   120
         TabIndex        =   161
         Top             =   5040
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   3836
         _Version        =   393216
         AllowUpdate     =   0   'False
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
         ColumnCount     =   9
         BeginProperty Column00 
            DataField       =   "ApplyItem"
            Caption         =   "Apply"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "Yes"
               FalseValue      =   "No"
               NullValue       =   "N/A"
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Description"
            Caption         =   "Description"
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
            DataField       =   "ItemAmount"
            Caption         =   "Amount"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """$""#,##0.00;(""$""#,##0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "ItemPercent"
            Caption         =   "Percent"
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
            DataField       =   "Basis"
            Caption         =   "Basis"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """$""#,##0.00;(""$""#,##0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "WageLow"
            Caption         =   "Low"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """$""#,##0.00;(""$""#,##0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "WageHigh"
            Caption         =   "High"
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
         BeginProperty Column07 
            DataField       =   "YTDMax"
            Caption         =   "Max"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """$""#,##0.00;(""$""#,##0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "Account"
            Caption         =   "Account"
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   659.906
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   929.764
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   929.764
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   929.764
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   929.764
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   929.764
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   929.764
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   929.764
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame6 
         Height          =   1815
         Left            =   7200
         TabIndex        =   162
         Top             =   3120
         Width           =   2415
         Begin VB.TextBox txtCheckDetail 
            DataField       =   "YTDGROSS"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$""#,##0.00;(""$""#,##0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   2
            EndProperty
            Height          =   285
            Index           =   36
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   167
            Text            =   "$"
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txtCheckDetail 
            DataField       =   "YTDFICA"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$""#,##0.00;(""$""#,##0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   2
            EndProperty
            Height          =   285
            Index           =   35
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   166
            Text            =   "$"
            Top             =   480
            Width           =   1095
         End
         Begin VB.TextBox txtCheckDetail 
            DataField       =   "YTDFIT"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$""#,##0.00;(""$""#,##0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   2
            EndProperty
            Height          =   285
            Index           =   34
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   165
            Text            =   "$"
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txtCheckDetail 
            DataField       =   "YTDSTATETAX"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$""#,##0.00;(""$""#,##0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   2
            EndProperty
            Height          =   285
            Index           =   33
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   164
            Text            =   "$"
            Top             =   960
            Width           =   1095
         End
         Begin VB.TextBox txtCheckDetail 
            DataField       =   "YTDLOCAL"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$""#,##0.00;(""$""#,##0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   2
            EndProperty
            Height          =   285
            Index           =   32
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   163
            Text            =   "$"
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "YTD Gross:"
            Height          =   255
            Index           =   30
            Left            =   0
            TabIndex        =   172
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "YTD FICA:"
            Height          =   255
            Index           =   29
            Left            =   0
            TabIndex        =   171
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "YTD FIT:"
            Height          =   255
            Index           =   28
            Left            =   0
            TabIndex        =   170
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "YTD State:"
            Height          =   255
            Index           =   27
            Left            =   0
            TabIndex        =   169
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "YTD Local:"
            Height          =   255
            Index           =   26
            Left            =   0
            TabIndex        =   168
            Top             =   1200
            Width           =   975
         End
      End
   End
   Begin VB.Frame frPayrollItems 
      Caption         =   "Payroll Items"
      Height          =   6255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   9735
      Begin VB.CommandButton cmdClosePyrlItem 
         Caption         =   "Back"
         Height          =   375
         Left            =   7920
         TabIndex        =   58
         Top             =   1080
         Width           =   1695
      End
      Begin VB.ComboBox cbPyrllItems 
         DataField       =   "Default"
         Height          =   315
         Index           =   2
         ItemData        =   "frm_Pay_Employees.frx":222E
         Left            =   7920
         List            =   "frm_Pay_Employees.frx":2265
         TabIndex        =   57
         Top             =   1800
         Width           =   1695
      End
      Begin VB.CommandButton cmdPyrlUpdate 
         Caption         =   "Update"
         Height          =   375
         Left            =   7920
         TabIndex        =   56
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdPyrlRefresh 
         Caption         =   "Refresh"
         Height          =   375
         Left            =   7920
         TabIndex        =   55
         Top             =   660
         Width           =   1695
      End
      Begin VB.CheckBox chkPyrllItem 
         Caption         =   "Employer"
         DataField       =   "EmployerYN"
         Height          =   255
         Left            =   240
         TabIndex        =   40
         Top             =   3960
         Width           =   975
      End
      Begin VB.Frame frAccount 
         Height          =   2175
         Left            =   120
         TabIndex        =   27
         Top             =   3960
         Width           =   9495
         Begin VB.CommandButton cmdAcct 
            Height          =   270
            Index           =   22
            Left            =   6960
            Picture         =   "frm_Pay_Employees.frx":238F
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   1320
            Width           =   375
         End
         Begin VB.CommandButton cmdAcct 
            Height          =   270
            Index           =   21
            Left            =   6960
            Picture         =   "frm_Pay_Employees.frx":24D9
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox txtPyrllItems 
            DataField       =   "Account3"
            Height          =   285
            Index           =   22
            Left            =   5280
            TabIndex        =   34
            Top             =   1320
            Width           =   1695
         End
         Begin VB.TextBox txtPyrllItems 
            DataField       =   "Account2"
            Height          =   285
            Index           =   21
            Left            =   5280
            TabIndex        =   32
            Top             =   480
            Width           =   1695
         End
         Begin VB.TextBox txtPyrllItems 
            DataField       =   "EmployerItemPercent"
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
            Index           =   10
            Left            =   1680
            TabIndex        =   30
            Top             =   960
            Width           =   1695
         End
         Begin VB.TextBox txtPyrllItems 
            DataField       =   "EmployerItemAmount"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$""#,##0.00;(""$""#,##0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   2
            EndProperty
            Height          =   285
            Index           =   9
            Left            =   1680
            TabIndex        =   28
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label lblAccts 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   22
            Left            =   3840
            TabIndex        =   37
            Top             =   1620
            Width           =   3495
         End
         Begin VB.Label lblAccts 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   21
            Left            =   3840
            TabIndex        =   36
            Top             =   780
            Width           =   3495
         End
         Begin VB.Label lblPyrllItems 
            Caption         =   "Credit Account:"
            Height          =   255
            Index           =   22
            Left            =   3840
            TabIndex        =   35
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label lblPyrllItems 
            Caption         =   "Debit Account:"
            Height          =   255
            Index           =   21
            Left            =   3840
            TabIndex        =   33
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label lblPyrllItems 
            Caption         =   "Percent of Basis:"
            Height          =   255
            Index           =   13
            Left            =   240
            TabIndex        =   31
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label lblPyrllItems 
            Caption         =   "Amount:"
            Height          =   255
            Index           =   12
            Left            =   240
            TabIndex        =   29
            Top             =   480
            Width           =   1335
         End
      End
      Begin VB.CommandButton cmdItemID 
         Height          =   270
         Left            =   3000
         Picture         =   "frm_Pay_Employees.frx":2623
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   600
         Width           =   375
      End
      Begin VB.Frame frEmployee 
         Caption         =   "Employee"
         Height          =   1695
         Left            =   120
         TabIndex        =   17
         Top             =   2160
         Width           =   9495
         Begin VB.CommandButton cmdAcct 
            Height          =   270
            Index           =   20
            Left            =   6960
            Picture         =   "frm_Pay_Employees.frx":276D
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox txtPyrllItems 
            DataField       =   "Account"
            Height          =   285
            Index           =   20
            Left            =   5280
            TabIndex        =   22
            Top             =   480
            Width           =   1695
         End
         Begin VB.TextBox txtPyrllItems 
            DataField       =   "ItemPercent"
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
            Index           =   7
            Left            =   1680
            TabIndex        =   20
            Top             =   960
            Width           =   1695
         End
         Begin VB.TextBox txtPyrllItems 
            DataField       =   "ItemAmount"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$""#,##0.00;(""$""#,##0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   2
            EndProperty
            Height          =   285
            Index           =   6
            Left            =   1680
            TabIndex        =   18
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label lblAccts 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   20
            Left            =   3840
            TabIndex        =   26
            Top             =   780
            Width           =   3495
         End
         Begin VB.Label lblPyrllItems 
            Caption         =   "Account:"
            Height          =   255
            Index           =   20
            Left            =   3840
            TabIndex        =   23
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label lblPyrllItems 
            Caption         =   "Percent of Basis:"
            Height          =   255
            Index           =   7
            Left            =   240
            TabIndex        =   21
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label lblPyrllItems 
            Caption         =   "Amount:"
            Height          =   255
            Index           =   6
            Left            =   240
            TabIndex        =   19
            Top             =   480
            Width           =   1335
         End
      End
      Begin VB.TextBox txtPyrllItems 
         Height          =   285
         Index           =   5
         Left            =   1320
         TabIndex        =   15
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox txtPyrllItems 
         DataField       =   "YTDMax"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00;(""$""#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   5280
         TabIndex        =   13
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox txtPyrllItems 
         DataField       =   "WageHigh"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00;(""$""#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   5280
         TabIndex        =   11
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox txtPyrllItems 
         DataField       =   "WageLow"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00;(""$""#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   5280
         TabIndex        =   9
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox txtPyrllItems 
         DataField       =   "Description"
         Height          =   285
         Index           =   1
         Left            =   5280
         TabIndex        =   7
         Top             =   600
         Width           =   1695
      End
      Begin VB.ComboBox cbPyrllItems 
         DataField       =   "Basis"
         Height          =   315
         Index           =   1
         ItemData        =   "frm_Pay_Employees.frx":28B7
         Left            =   1320
         List            =   "frm_Pay_Employees.frx":28C4
         TabIndex        =   4
         Top             =   1320
         Width           =   2055
      End
      Begin VB.ComboBox cbPyrllItems 
         DataField       =   "Type"
         Height          =   315
         Index           =   0
         ItemData        =   "frm_Pay_Employees.frx":28D9
         Left            =   1320
         List            =   "frm_Pay_Employees.frx":28E9
         TabIndex        =   3
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox txtPyrllItems 
         DataField       =   "PyrlItemID"
         Height          =   285
         Index           =   0
         Left            =   1320
         TabIndex        =   1
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox tempPayItems 
         Height          =   285
         Left            =   1800
         TabIndex        =   181
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtPayRollItems 
         DataField       =   "W2Box14"
         Height          =   285
         Index           =   14
         Left            =   8880
         TabIndex        =   194
         Top             =   4200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtPayRollItems 
         DataField       =   "W2Box13b"
         Height          =   285
         Index           =   13
         Left            =   8280
         TabIndex        =   195
         Top             =   4200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtPayRollItems 
         DataField       =   "W2Box13"
         Height          =   285
         Index           =   12
         Left            =   7680
         TabIndex        =   196
         Top             =   4200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtPayRollItems 
         DataField       =   "W2Box15"
         Height          =   285
         Index           =   15
         Left            =   7680
         TabIndex        =   197
         Top             =   4440
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtPayRollItems 
         DataField       =   "W2Box17"
         Height          =   285
         Index           =   16
         Left            =   8280
         TabIndex        =   198
         Top             =   4440
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtPayRollItems 
         DataField       =   "W2Box18"
         Height          =   285
         Index           =   17
         Left            =   8880
         TabIndex        =   199
         Top             =   4440
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtPayRollItems 
         DataField       =   "W2Box21"
         Height          =   285
         Index           =   19
         Left            =   8280
         TabIndex        =   200
         Top             =   4680
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtPayRollItems 
         DataField       =   "W2Box20"
         Height          =   285
         Index           =   18
         Left            =   7680
         TabIndex        =   201
         Top             =   4680
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtPayRollItems 
         DataField       =   "Frm941Line2"
         Height          =   285
         Index           =   20
         Left            =   7680
         TabIndex        =   202
         Top             =   4920
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtPayRollItems 
         DataField       =   "Frm941Line3"
         Height          =   285
         Index           =   21
         Left            =   8280
         TabIndex        =   203
         Top             =   4920
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtPayRollItems 
         DataField       =   "Frm941Line4"
         Height          =   285
         Index           =   22
         Left            =   8880
         TabIndex        =   204
         Top             =   4920
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtPayRollItems 
         DataField       =   "Frm941Line6a"
         Height          =   285
         Index           =   23
         Left            =   7680
         TabIndex        =   205
         Top             =   5160
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtPayRollItems 
         DataField       =   "Frm941Line6b"
         Height          =   285
         Index           =   24
         Left            =   8280
         TabIndex        =   206
         Top             =   5160
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtPayRollItems 
         DataField       =   "Frm941Line7"
         Height          =   285
         Index           =   25
         Left            =   8880
         TabIndex        =   207
         Top             =   5160
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtPayRollItems 
         DataField       =   "Frm941Line9"
         Height          =   285
         Index           =   26
         Left            =   7680
         TabIndex        =   208
         Top             =   5400
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtPayRollItems 
         DataField       =   "Frm941Line12"
         Height          =   285
         Index           =   27
         Left            =   8280
         TabIndex        =   209
         Top             =   5400
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtPayRollItems 
         DataField       =   "W2Box12"
         Height          =   285
         Index           =   11
         Left            =   8880
         TabIndex        =   182
         Top             =   3120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtPayRollItems 
         DataField       =   "W2Box11"
         Height          =   285
         Index           =   10
         Left            =   8280
         TabIndex        =   183
         Top             =   3120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtPayRollItems 
         DataField       =   "W2Box10"
         Height          =   285
         Index           =   9
         Left            =   7680
         TabIndex        =   184
         Top             =   3120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtPayRollItems 
         DataField       =   "W2Box9"
         Height          =   285
         Index           =   8
         Left            =   8880
         TabIndex        =   185
         Top             =   2880
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtPayRollItems 
         DataField       =   "W2Box8"
         Height          =   285
         Index           =   7
         Left            =   8280
         TabIndex        =   186
         Top             =   2880
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtPayRollItems 
         DataField       =   "W2Box7"
         Height          =   285
         Index           =   6
         Left            =   7680
         TabIndex        =   187
         Top             =   2880
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtPayRollItems 
         DataField       =   "W2Box6"
         Height          =   285
         Index           =   5
         Left            =   8880
         TabIndex        =   188
         Top             =   2640
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtPayRollItems 
         DataField       =   "W2Box5"
         Height          =   285
         Index           =   4
         Left            =   8280
         TabIndex        =   189
         Top             =   2640
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtPayRollItems 
         DataField       =   "W2Box4"
         Height          =   285
         Index           =   3
         Left            =   7680
         TabIndex        =   190
         Top             =   2640
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtPayRollItems 
         DataField       =   "W2Box3"
         Height          =   285
         Index           =   2
         Left            =   8880
         TabIndex        =   191
         Top             =   2400
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtPayRollItems 
         DataField       =   "W2Box2"
         Height          =   285
         Index           =   1
         Left            =   8280
         TabIndex        =   192
         Top             =   2400
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtPayRollItems 
         DataField       =   "W2Box1"
         Height          =   285
         Index           =   0
         Left            =   7680
         TabIndex        =   193
         Top             =   2400
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblPyrllItems 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Item Tracking:"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   8
         Left            =   7920
         TabIndex        =   59
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label lblPyrllItems 
         Caption         =   "Desc:"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   16
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label lblPyrllItems 
         Caption         =   "Maximum Annual:"
         Height          =   255
         Index           =   4
         Left            =   3840
         TabIndex        =   14
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label lblPyrllItems 
         Caption         =   "YTD Gross High:"
         Height          =   255
         Index           =   3
         Left            =   3840
         TabIndex        =   12
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label lblPyrllItems 
         Caption         =   "YTD Gross Low:"
         Height          =   255
         Index           =   2
         Left            =   3840
         TabIndex        =   10
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lblPyrllItems 
         Caption         =   "Description:"
         Height          =   255
         Index           =   1
         Left            =   3840
         TabIndex        =   8
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lblPyrllItems 
         Caption         =   "Basis:"
         Height          =   255
         Index           =   24
         Left            =   240
         TabIndex        =   6
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label lblPyrllItems 
         Caption         =   "Type:"
         Height          =   255
         Index           =   23
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lblPyrllItems 
         Caption         =   "Item ID:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   975
      End
   End
End
Attribute VB_Name = "frm_Pay_Employees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ADOprimaryrs As ADODB.Recordset
Attribute ADOprimaryrs.VB_VarHelpID = -1
Dim ADOCreatePay As ADODB.Recordset
Dim rsgrd As ADODB.Recordset
Dim Rst As ADODB.Recordset
Dim rstPyrl As ADODB.Recordset
Dim rstSETUP As ADODB.Recordset
'Dim rstFIT As ADODB.Recordset
Dim rsCommissions As ADODB.Recordset
Dim rstItems As ADODB.Recordset
Dim rscheckGrd As ADODB.Recordset
Dim rsCheck As ADODB.Recordset
Dim rsPyrlSetup As ADODB.Recordset

Dim db As ADODB.Connection

Dim Numrec, K, Regular, OT, GROSS, PRETAXDED, NETADDITIONS, AGI, FICAER, FicaEE, FUTA, SUI, SocSec, Medi, FIT, STATETAX, LocalTax, Net, AllowAmt, Periods, ADDITIONS, DEDUCTIONS, GROSSDED, AGIDED, LASTHOURS, Commission, ItemTotalAmount, MaxValue, OTConversion

Dim Criteria As String, strSQL As String, LoadSelect As Boolean

Dim NextCheck$
Private Sub LoadPayrollDB()
ShowStatus True
'If the type PyrlItemID exists then load that record
'Otherwise create a new record w/ this PyrlItemID.
        
    If ADOCreatePay Is Nothing Then
        Set ADOCreatePay = New ADODB.Recordset
    Else
        ADOCreatePay.CancelUpdate
        ADOCreatePay.Close
        Set ADOCreatePay = Nothing
        Set ADOCreatePay = New ADODB.Recordset
    End If
    
    If txtPyrllItems(0) <> "" Then
        ADOCreatePay.Open "SELECT * FROM [Pyrl - Payroll Items] WHERE [PyrlItemID]='" & txtPyrllItems(0) & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
    Else
        ADOCreatePay.Open "SELECT * FROM [Pyrl - Payroll Items]", db, adOpenKeyset, adLockOptimistic, adCmdText
    End If
    
    cmdPyrlUpdate.Enabled = True
    cmdPyrlRefresh.Enabled = True
    chkPyrllItem.Enabled = True
        
    If CheckNewDB(ADOCreatePay, "Payroll Items") = True Then
        tempPayItems = txtPyrllItems(0).Text
        NewLoadPayroll
    End If
    
    Dim textPay As TextBox
    For Each textPay In Me.txtPyrllItems
        Set textPay.DataSource = ADOCreatePay
        Select Case textPay.Index
        Case 20, 21, 22
           If textPay.Text <> "" Then
                lblAccts(textPay.Index) = LookRecord("[GL COA Account Name]", "[GL Chart Of Accounts]", db, "[GL COA Account No] = '" & textPay.Text & "'")
           End If
        End Select
    Next
    
    Dim textPayRoll As TextBox
    For Each textPayRoll In Me.txtPayRollItems
        Set textPayRoll.DataSource = ADOCreatePay
    Next
    
    Set chkPyrllItem.DataSource = ADOCreatePay
    
    Dim cbPay As ComboBox
    For Each cbPay In Me.cbPyrllItems
        Set cbPay.DataSource = ADOCreatePay
    Next
    
    frEmployee.Enabled = True
        If ADOCreatePay![EmployerYN] = True Then
            frAccount.Enabled = True
        Else
            frAccount.Enabled = False
        End If
    
TypeDesc

If tempPayItems = "" Then
    tempPayItems = txtPyrllItems(0).Text
Else
    If IsNull(ADOCreatePay![PyrlItemID]) Then
        txtPyrllItems(0).Text = tempPayItems
    Else
        tempPayItems = ADOCreatePay![PyrlItemID]
    End If
End If
ShowStatus False

End Sub

Private Sub NewLoadPayroll()
      ADOCreatePay.AddNew
      cmdPyrlRefresh.Enabled = False
      cmdItemID.Enabled = False
      cmdClosePyrlItem.Caption = "Cancel"
End Sub


Private Sub ClearValues()
Dim i As Integer
    For i = 0 To txtPayRollItems.UBound
        txtPayRollItems(i).Text = ""
    Next
End Sub

Private Sub cbPyrllItems_Click(Index As Integer)
If Index = 2 Then
Select Case cbPyrllItems(2).Text
    Case "User Define or None"
        ClearValues
    
    Case "Compensation"
        ClearValues
        txtPayRollItems(0) = "Add To"
        txtPayRollItems(20) = "Add To"
        
    Case "Non-Taxable Sick Pay"
        ClearValues
        txtPayRollItems(12) = "Add To"
        txtPayRollItems(13) = "J"
    
    Case "Advanced EIC"
        ClearValues
        txtPayRollItems(8) = "Add To"
        txtPayRollItems(11) = "Add To"
    
    Case "401K Plan"
        ClearValues
        txtPayRollItems(0) = "Subtract From"
        txtPayRollItems(12) = "Add To"
        txtPayRollItems(13) = "D"
        txtPayRollItems(15) = "3"
        txtPayRollItems(20) = "Subtract From"
        
    Case "Other"
        ClearValues
        txtPayRollItems(14) = "Add To"
        
    
    
    Case "Tips"
        ClearValues
        txtPayRollItems(0) = "Add To"
        txtPayRollItems(6) = "Add To"
        txtPayRollItems(20) = "Add To"
        txtPayRollItems(24) = "Add To"
        txtPayRollItems(23) = "Subtract From"
    
    Case "Qualified Moving Expense"
        ClearValues
        txtPayRollItems(12) = "Add To"
        txtPayRollItems(13) = "P"
    
    
    Case "Other Moving Expense"
        ClearValues
        txtPayRollItems(0) = "Add To"
        txtPayRollItems(11) = "Add To"
        txtPayRollItems(14) = "Add To"
        txtPayRollItems(20) = "Add To"
    
    Case "State Income Tax"
        ClearValues
        txtPayRollItems(17) = "Add To"
    
    Case "Local Income Tax"
        ClearValues
        txtPayRollItems(19) = "Add To"
    
    Case "403(b) Plan"
        ClearValues
        txtPayRollItems(0) = "Subtract From"
        txtPayRollItems(12) = "Add To"
        txtPayRollItems(13) = "E"
        txtPayRollItems(15) = "3"
        txtPayRollItems(20) = "Subtract From"
    
    Case "408(k)(6) SEP Plan"
        ClearValues
        txtPayRollItems(0) = "Subtract From"
        txtPayRollItems(12) = "Add To"
        txtPayRollItems(13) = "F"
        txtPayRollItems(15) = "3"
        txtPayRollItems(20) = "Subtract From"
    
    Case "Elective 457(b) Plan"
        ClearValues
        txtPayRollItems(0) = "Subtract From"
        txtPayRollItems(12) = "Add To"
        txtPayRollItems(13) = "G"
        txtPayRollItems(15) = "3"
        txtPayRollItems(20) = "Subtract From"
    
    Case "501(c)(18))d) Plan"
        ClearValues
        txtPayRollItems(12) = "Add To"
        txtPayRollItems(13) = "H"
    
    Case "SEC 457 Distribution"
        ClearValues
        txtPayRollItems(0) = "Add To"
        txtPayRollItems(10) = "Add To"
        txtPayRollItems(20) = "Add To"
    
    Case "Fringe Benefits"
        ClearValues
        txtPayRollItems(0) = "Add To"
        txtPayRollItems(11) = "Add To"
        txtPayRollItems(20) = "Add To"
End Select
End If

TypeDesc

End Sub

Private Sub cbPyrllItems_KeyPress(Index As Integer, KeyAscii As Integer)
Dim keyResponse As Boolean
    keyResponse = CtrlValidate(KeyAscii, "")
    If keyResponse = True Then
    Else
       KeyAscii = 0
    End If
End Sub

Private Sub cbPyrllItems_LostFocus(Index As Integer)
    If CbValidate(cbPyrllItems(Index), cbPyrllItems(Index).Text) = False Then
       MsgBox "There is no such selection", vbInformation, "Information"
    End If
End Sub


Private Sub cmdAcct_Click(Index As Integer)
    Dim No As Integer
    Dim sql As String
    Dim ghead As String
    Dim fhead As String
    
    No = 1600
    sql = "select [GL COA Account No], [GL COA Account Name]" & _
        "from [GL Chart of Accounts]"
    ghead = "Account Description"
    fhead = "Account No//Description"
    
    AllLookup.GetWhichTable No, sql, ghead, fhead, db
    'AllLookup.Show vbModal
    lblAccts(Index).Caption = lblAcctTemp.Caption
    txtPyrllItems(Index).Text = txtFieldsTemp.Text
    txtPyrllItems(Index).SetFocus
End Sub

Private Sub cmdBackCheck_Click()
    Me.Height = 6645
    Me.Width = 9825
    frCheckDetail.Visible = False
    If rscheckGrd Is Nothing Then
    Else
        rscheckGrd.Close
        Set rscheckGrd = Nothing
        rsCheck.Close
        Set rsCheck = Nothing
    End If
    Me.Caption = "Pay Employee"
End Sub

Private Sub cmdCheckEmplookup_Click()
On Error Resume Next          '---------on error statement

Dim TempData As String
Static LoadAlready As Boolean

    Dim SQLstatement As String
    Dim No As Integer
    Dim ghead As String
    Dim fhead As String
    
    TempData = txtCheckDetail(0)
    
    No = 1700
    SQLstatement = "select [EMP ID], [EMP Name]" & _
                    "from [EMP Employees]"
    ghead = "Employee"
    fhead = "ID//Name"
    
    AllLookup.GetWhichTable No, SQLstatement, ghead, fhead, db
    'AllLookup.Show vbModal
    
    If TempData = txtCheckDetail(0) Then Exit Sub
    TempEMPID = txtCheckDetail(0)
    lblName = LookRecord("[EMP Name]", "[EMP Employees]", db, "[EMP ID] = '" & TempEMPID & "'")
    
    BlankCheck
    ResetPyrlItems
    Set DataGrid1.DataSource = Nothing
    If rscheckGrd Is Nothing Then
    Else
        rscheckGrd.CancelUpdate
        rscheckGrd.Close
        Set rscheckGrd = Nothing
    End If
        Set rscheckGrd = New ADODB.Recordset
        rscheckGrd.Open "Select * from [Pyrl - Select Pyrl Items Work]", db, adOpenKeyset, adLockOptimistic, adCmdText
        'Debug.Print "Select * from [Pyrl - Register Detail] WHERE [EMP ID]='" & txtCheckDetail(0) & "'"
        Set DataGrid1.DataSource = rscheckGrd
    Set DataGrid1.DataSource = rscheckGrd
    
    If rsCheck Is Nothing Then
    Else
        rsCheck.CancelUpdate
        rsCheck.Close
        Set rsCheck = Nothing
    End If
    Set rsCheck = New ADODB.Recordset
    rsCheck.Open "Select * from [Pyrl - Employees] WHERE [Emp ID]='" & txtCheckDetail(0) & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
 
    Dim Ctrltxt As TextBox
    For Each Ctrltxt In Me.txtCheckDetail
         Set Ctrltxt.DataSource = rsCheck
         If rsCheck("" & Ctrltxt.DataField & "").Type = 202 Then Ctrltxt.MaxLength = rsCheck("" & Ctrltxt.DataField & "").DefinedSize
    Next
    txtCheckDetail(30).Text = "$0.00"
    txtCheckDetail(31).Text = "$0.00"
End Sub

Private Sub CmdClearComm_Click()
  With rsgrd
     .MoveFirst
        Do Until .EOF
        'rst '.Edit
         ![AR SALE Select to Pay] = 0
         .Update
         .MoveNext
        Loop
  End With
End Sub

Private Sub cmdClose_Click()
    Set dtgrdDataGrid.DataSource = Nothing
    rsgrd.Close
    Set rsgrd = Nothing
    frCommision.Visible = False
    Form_Resize
    Me.Caption = "Pay Employees"
End Sub

Private Sub cmdClosePyrlItem_Click()
If cmdClosePyrlItem.Caption = "Back" Then
    If txtPyrllItems(0) <> "" Then         '-enable this when finish
        ADOCreatePay.CancelUpdate
        ADOCreatePay.Close
    End If
    
    Set ADOCreatePay = Nothing
    
    Dim textPayRoll As TextBox
    For Each textPayRoll In Me.txtPayRollItems
        Set textPayRoll.DataSource = Nothing
    Next
    
    Dim EmptyText As TextBox
    For Each EmptyText In Me.txtPyrllItems
        Set EmptyText.DataSource = Nothing
        EmptyText = ""
    Next
    Dim Emptycb As ComboBox
    For Each Emptycb In Me.cbPyrllItems
        Set Emptycb.DataSource = Nothing
        Emptycb = ""
    Next
    frPayrollItems.Top = 0
    frPayrollItems.Left = 0
    Me.Caption = "Pay Employee"
    frPayrollItems.Visible = False
    Form_Resize
    cmdPyrlUpdate.Enabled = False
    cmdPyrlRefresh.Enabled = False
    chkPyrllItem.Enabled = False
Else
    ADOCreatePay.CancelUpdate
    'ADOCreatePay.MoveFirst
    txtPyrllItems(0).Text = ""
    LoadPayrollDB
    cmdClosePyrlItem.Caption = "Back"
    cmdPyrlRefresh.Enabled = True
    cmdItemID.Enabled = True
End If
End Sub

Private Sub cmdCreatePayrollChk_Click()
ShowStatus True
If txtCheckDetail(0) = "" Then
    MsgBox "Please select an employee first", vbInformation, "Error"
    Exit Sub
End If
If MsgBox("Are You Sure?", vbYesNo) = vbNo Then
    ShowStatus False
    Exit Sub
End If

PostToRegister

Dim rsCPCheckDeatail As ADODB.Recordset                         'recalc total checks to print
'Dim d As Database
'Set d = CurrentDb
Set rsCPCheckDeatail = New ADODB.Recordset
rsCPCheckDeatail.Open "SELECT [SumofNetPay] FROM [Pyrl - Total Checks]", db, adOpenKeyset, adLockOptimistic, adCmdTable

If rsCPCheckDeatail.RecordCount > 0 Then
    lblCheckPrints.Caption = FormatCurr(NZ(rsCPCheckDeatail!SumofNetPay, 0))
Else
    lblCheckPrints.Caption = "$0.00"
End If

rsCPCheckDeatail.Close
Set rsCPCheckDeatail = Nothing

'If Me!cmdZoom.Visible = True Then
'    Call EMP_ID_AfterUpdate
'End If
ShowStatus False
End Sub

Private Sub cmdDateComm_Click()
    Menu_Calendar.WhoCallMe True, 1545
    'Menu_Calendar.Show vbModal
End Sub


Private Sub cmdEmpLookup_Click()
    Dim SQLstatement As String
    Dim No As Integer
    Dim ghead As String
    Dim fhead As String

    No = 1590
    SQLstatement = "select [EMP ID], [EMP Name]" & _
                    "from [EMP Employees]"
    ghead = "Employee"
    fhead = "ID//Name"
    
    AllLookup.GetWhichTable No, SQLstatement, ghead, fhead, db
    'AllLookup.Show vbModal
    If txtCommission(0) <> "" Then
        TempEMPID = txtCommission(0)
        lblName = LookRecord("[EMP Name]", "[EMP Employees]", db, "[EMP ID] = '" & TempEMPID & "'")
        With rsgrd
           .MoveFirst
              Do Until .EOF
               'rst '.Edit
               If ![EMP ID] = txtCommission(0) Then ![AR SALE Select to Pay] = -1
               .Update
               .MoveNext
              Loop
        End With
    End If
End Sub

Private Sub TypeDesc()
Dim txtmessage As String

Select Case cbPyrllItems(0)
Case "Addition"
    lblPyrllItems(20).Caption = "Debit Account"
   txtmessage = ""
   Select Case cbPyrllItems(1)
        Case "Gross"
            txtmessage = "Taxable Income"
        Case "AGI"
            txtmessage = "NonTaxable Income"
        Case "Net"
             txtmessage = "Reimbursement"
    End Select

Case "Deduction"
    lblPyrllItems(20).Caption = "Credit Account"
     txtmessage = ""
    Select Case cbPyrllItems(1)
        Case "Gross"
            txtmessage = "NonTaxable Deduction"
        Case "AGI"
            txtmessage = "AfterTax Deduction"
        Case "Net"
            txtmessage = "AfterTax Deduction"
     End Select

Case "State Tax"
    lblPyrllItems(20).Caption = "Credit Account"
    txtmessage = "State Tax"
    

Case "Local Tax"
    lblPyrllItems(20).Caption = "Credit Account"
    txtmessage = "Local Tax"
    

Case ""
    lblPyrllItems(20).Caption = "Account"
    txtmessage = ""
End Select
    txtPyrllItems(5) = txtmessage
End Sub

Private Sub cmdItemID_Click()
    Dim TempStr As String
    Dim SQLstatement As String
    Dim No As Integer
    Dim ghead As String
    Dim fhead As String
    
    TempStr = txtPyrllItems(0)
    
    No = 1610
    SQLstatement = "select [PyrlItemID], [Description]" & _
                    "from [Pyrl - Payroll Items]"
    ghead = "Payroll Items"
    fhead = "ID//Description"
    
    AllLookup.GetWhichTable No, SQLstatement, ghead, fhead, db
    'AllLookup.Show vbModal
    txtPyrllItems(0).SetFocus
    tempPayItems = txtPyrllItems(0).Text
    If TempStr <> txtPyrllItems(0) Then LoadPayrollDB
    
End Sub

Private Sub cmdPyrlUpdate_Click()
Dim i As Integer
    
    For i = 0 To 23
    Select Case i
        Case 0, 1, 2
            If Trim(txtPyrllItems(i)) = "" Then
                MsgBox lblPyrllItems(i) & " is empty. Please fill it", vbInformation, "Information"
                Exit Sub
            End If
            If Trim(cbPyrllItems(i)) = "" Then
                MsgBox "One of the ComboBox is empty. Please make your selection", vbInformation, "Information"
                Exit Sub
            End If
        Case 3, 4, 5, 6, 7, 20
            If Trim(txtPyrllItems(i)) = "" Then
                MsgBox lblPyrllItems(i) & " is empty. Please fill it", vbInformation, "Information"
                Exit Sub
            End If
    End Select
    Next
    ADOCreatePay.Update
    cmdClosePyrlItem.Caption = "Back"
    cmdPyrlRefresh.Enabled = True
    cmdItemID.Enabled = True
End Sub

Private Sub cmdSSelectAll_Click()
  With rsgrd
     .MoveFirst
        Do Until .EOF
         'rst '.Edit
         ![AR SALE Select to Pay] = -1
         .Update
         .MoveNext
        Loop
  End With
End Sub

Private Sub Combo1_Click()
  ADOprimaryrs.CancelUpdate
  ADOprimaryrs.Close
  OpenDB "select [PAY],[EMP ID],[LASTHOURS],[NETCHECK],[PAYFREQUENCY],[PAYTYPE],[LASTDATE] from [Pyrl - Employee Data]"
End Sub

Private Sub Combo2_Click()
  ADOprimaryrs.Close
  OpenDB "select [PAY],[EMP ID],[LASTHOURS],[NETCHECK],[PAYFREQUENCY],[PAYTYPE],[LASTDATE] from [Pyrl - Employee Data]"
End Sub



Private Sub DataGrid1_AfterColEdit(ByVal ColIndex As Integer)
      SendKeys ("{ENTER}")
  If grdDataGrid.Row > 0 Then
      SendKeys ("{up}")
      SendKeys ("{down}")
  ElseIf grdDataGrid.Row = 0 Then
      SendKeys ("{down}")
      SendKeys ("{up}")
  End If
End Sub

Private Sub DataGrid1_ButtonClick(ByVal ColIndex As Integer)
If DataGrid1.Row = -1 Then Exit Sub
Select Case ColIndex
Case 0
         SendKeys ("{ENTER}")
   If DataGrid1.Columns(0).Text = "No" Then
      DataGrid1.Columns(0).Text = "Yes"
   Else
      DataGrid1.Columns(0).Text = "No"
   End If
         SendKeys ("{ENTER}")
         SendKeys ("{down}")
         SendKeys ("{up}")

End Select

CalcPayroll

End Sub

Private Sub DataGrid1_GotFocus()
    DataGrid1.Columns(0).Button = True
End Sub


Private Sub dtgrdDataGrid_AfterColEdit(ByVal ColIndex As Integer)
      SendKeys ("{ENTER}")
  If grdDataGrid.Row > 0 Then
      SendKeys ("{up}")
      SendKeys ("{down}")
  ElseIf grdDataGrid.Row = 0 Then
      SendKeys ("{down}")
      SendKeys ("{up}")
  End If
End Sub

Private Sub dtgrdDataGrid_ButtonClick(ByVal ColIndex As Integer)
If DataGrid1.Row = -1 Then Exit Sub
Select Case ColIndex
Case 0
         SendKeys ("{ENTER}")
   If dtgrdDataGrid.Columns(0).Text = "No" Then
      dtgrdDataGrid.Columns(0).Text = "Yes"
   Else
      dtgrdDataGrid.Columns(0).Text = "No"
   End If
         SendKeys ("{ENTER}")
         SendKeys ("{down}")
         SendKeys ("{up}")

End Select
End Sub


Private Sub dtgrdDataGrid_Error(ByVal DataError As Integer, Response As Integer)
    If DataGridKnownError(DataError) Then
        Response = 0
    End If
End Sub

Private Sub dtgrdDataGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If grdDataGrid.Row < 0 Then Exit Sub
     grdDataGrid.AllowUpdate = False
End Sub

Private Sub Form_Load()
On Error GoTo FormErr
ShowStatus True
  Set db = New ADODB.Connection
  db.CursorLocation = adUseClient
  db.Open gblADOProvider
  
    Me.Height = 6645
    Me.Width = 9825
  
  Combo1.Text = Combo1.List(4)
  Combo2.Text = Combo2.List(5)
  txtFieldsDate(0) = FormatDate(Date)

  OpenDB "select [PAY],[EMP ID],[LASTHOURS],[NETCHECK],[PAYFREQUENCY],[PAYTYPE],[LASTDATE] from [Pyrl - Employee Data]"
    
  grdDataGrid.Columns(0).Button = True
  grdDataGrid.Columns(6).Button = True
  
  GetTextColor Me
    
ShowStatus False
Exit Sub
FormErr:
  MsgBox Err.Description
  ShowStatus False
End Sub

Private Sub OpenDB(SQLstatement)
  
Set ADOprimaryrs = New ADODB.Recordset

If Combo1.Text = "All" And Combo2.Text = "All" Then
  ADOprimaryrs.Open SQLstatement, db, adOpenKeyset, adLockOptimistic, adCmdText
ElseIf Combo1.Text = "All" Then
  ADOprimaryrs.Open SQLstatement & "WHERE  [PAYTYPE]='" & Combo2.Text & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
ElseIf Combo2.Text = "All" Then
  ADOprimaryrs.Open SQLstatement & "WHERE [PAYFREQUENCY]='" & Combo1.Text & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
Else
  ADOprimaryrs.Open SQLstatement & "WHERE [PAYFREQUENCY]='" & Combo1.Text & "' AND [PAYTYPE]='" & Combo2.Text & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
End If

Set grdDataGrid.DataSource = Nothing

Set grdDataGrid.DataSource = ADOprimaryrs

'
CheckToPrint

End Sub

Private Sub CheckToPrint()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open "[Pyrl - Total Checks]", db, adOpenKeyset, adLockOptimistic, adCmdTable
    If rs.RecordCount = 0 Then
        lblCheckPrints.Caption = "$0.00"
    Else
        lblCheckPrints.Caption = FormatCurr(rs![SumofNetPay])
    End If
    rs.Close
    Set rs = Nothing
End Sub


Private Sub CmdClear_Click()
'Dim dbs As Database, rst As ADODB.Recordset
'Dim strMESSAGE$, strSQL$


On Error Resume Next
    

'    strMESSAGE = "Are You Sure You Want to Clear Marked Employees?"
'    If MsgBox(strMESSAGE, vbYesNo) = vbNo Then
'    Exit Sub
'    End If
    
'    Set dbs = CurrentDb
    
    'Open Employee Data recordset
'    strSQL$ = "SELECT * FROM [Pyrl - Employee Data] where [pay] = -1"
'    Set rst = dbs.OpenRecordset(strSQL$)
'    If rst.RecordCount > 0 Then
  With ADOprimaryrs
     .MoveFirst
        Do Until .EOF
        'rst '.Edit
         ![PAY] = 0
         ![NETCHECK] = 0
         .Update
         .MoveNext
        Loop
  End With
    
'    rst.MoveFirst
'        Do Until rst.EOF
        'rst '.Edit
'        rst![PAY] = 0
'        rst!NETCHECK = 0
'        rst.Update
'        rst.MoveNext
'        Loop
'    End If
'    rst.Close
'    Set dbs = Nothing
'    Me![Pyrl - Employee Data subform1].Form.Requery
  
'  Exit Sub
'CmdClear_Click_Error:
'  Call ErrorLog("Pay Employees", "CmdClear_Click", Now,  Err.number, Err.description, True,db)
'  Resume Next

End Sub

Private Sub cmdCommissions_Click()
'On Error GoTo Err_cmdCommissions_Click
If CheckEmpty = True Then Exit Sub
        
    txtCommission(1) = txtFieldsDate(1)
    txtCommission(2) = txtFieldsDate(2)
    'Dim stDocName As String
    'Dim stLinkCriteria As String

    Me.Caption = "Sales Commission"
    frCommision.Left = 0
    frCommision.Top = 0
    frCommision.Visible = True
    frCommision.ZOrder 0
    Form_Resize
    
    Dim SQLstatement As String
    
    Set rsgrd = New ADODB.Recordset
    SQLstatement = "select [EMP ID],[AR SALE Select to Pay],[EMP Name],[AR SALE Ext Document #],[AR SALE Document Type],[AR SALE Date],[COMMPERCENT],[Method],[Commission],[AR SALE Customer ID],[Price],[Cost],[Profit] from [Pyrl - Commission Select]"
    rsgrd.Open SQLstatement, db, adOpenKeyset, adLockOptimistic, adCmdText
    
    Set dtgrdDataGrid.DataSource = rsgrd
    
    'DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_cmdCommissions_Click:
    Exit Sub

Err_cmdCommissions_Click:
    MsgBox Err.Description
    Resume Exit_cmdCommissions_Click

End Sub

Private Sub RefreshDtGrd()
  
  Set grdDataGrid.DataSource = Nothing
  ADOprimaryrs.Close
  Set ADOprimaryrs = Nothing

  OpenDB "select [PAY],[EMP ID],[LASTHOURS],[NETCHECK],[PAYFREQUENCY],[PAYTYPE],[LASTDATE] from [Pyrl - Employee Data]"

End Sub

Private Function CheckEmpty(Optional Neglect As Boolean) As Boolean
If txtfields = "" Then
   MsgBox "Please Select Bank Account."
   CheckEmpty = True
   Exit Function
End If
If txtFieldsDate(1) = "" Then
   MsgBox "Enter a pay period starting date."
   CheckEmpty = True
   Exit Function
End If

If txtFieldsDate(2) = "" Then
   MsgBox "Enter a pay period ending date."
   CheckEmpty = True
   Exit Function
End If

If TempEMPID = "" And Neglect = False Then
   MsgBox "Please select employee first."
   CheckEmpty = True
   Exit Function
End If
   CheckEmpty = False
End Function

Private Sub cmdCreatePayroll_Click()
Dim rs As ADODB.Recordset

If CheckEmpty() = True Then Exit Sub

If MsgBox("Creating payroll for all selected employee." & vbCr & "Are You Sure?", vbYesNo) = vbNo Then
    ShowStatus False
    Exit Sub
End If

ShowStatus True

Call CalcPayroll
Call PostToRegister
'RefreshDtGrd

Set rs = New ADODB.Recordset
rs.Open "[Pyrl - Total Checks]", db, adOpenKeyset, adLockOptimistic, adCmdTable
If rs.RecordCount > 0 Then
    lblCheckPrints = FormatCurr(rs!SumofNetPay)
Else
    lblCheckPrints = "$0.00"
End If

rs.Close
Set rs = Nothing

LoadSelect = True
Set grdDataGrid.DataSource = Nothing
    ADOprimaryrs.Requery
Set grdDataGrid.DataSource = ADOprimaryrs
LoadSelect = False
ShowStatus False
End Sub

Private Sub cmdDate_Click(Index As Integer)
Select Case Index
Case 0
    Menu_Calendar.WhoCallMe True, 1500
Case 1
    Menu_Calendar.WhoCallMe True, 1510
Case 2
    Menu_Calendar.WhoCallMe True, 1520
End Select
    'Menu_Calendar.Show vbModal
End Sub

Private Sub cmdLookup_Click()
    Dim SQLstatement As String
    Dim No As Integer
    Dim ghead As String
    Dim fhead As String

    No = 1510
    SQLstatement = "select [BANK ACCT ID], [BANK ACCT Name]" & _
                    "from [BANK Accounts]"
    ghead = "Account Description"
    fhead = "Account No//Description"
    
    AllLookup.GetWhichTable No, SQLstatement, ghead, fhead, db
    'AllLookup.Show vbModal
    If txtfields <> "" Then
        lblAcctBal = NZ(LookRecord("[GL COA Account Balance]", "[GL Chart Of Accounts]", db, "[GL COA Account No] = '" & txtfields & "'"))
        lblAcctBal = FormatCurr(lblAcctBal)
    End If
End Sub

Private Sub cmdManualChecks_Click()
    
If CheckEmpty = True Then Exit Sub

'Pyrl - Register

frCheckDetail.Top = 0
frCheckDetail.Left = 0
frCheckDetail.ZOrder 0
frCheckDetail.Visible = True
Me.Caption = "Payroll - Check Detail"
Me.Height = 7710
    
    'DoCmd.OpenForm "Pyrl - Check Detail"
    
    'Forms![Pyrl - Check Detail]![ckCheckDetail].Value = False
    'Forms![Pyrl - Check Detail]!txtBlankCheckHourlyRate.Visible = True
    'Forms![Pyrl - Check Detail]!txtBlankCheckOTRate.Visible = True
    'Forms![Pyrl - Check Detail]!txtHourlyRate.Visible = False
    'Forms![Pyrl - Check Detail]!txtOTRate.Visible = False
    'Forms![Pyrl - Check Detail]!txtSalary.Visible = False
    'Forms![Pyrl - Check Detail]!txtCommission.Visible = False
    'Forms![Pyrl - Check Detail]!txtAmount.Visible = True
    
    'Forms![Pyrl - Check Detail]!txtLastFica.Locked = False
    'Forms![Pyrl - Check Detail]!txtLastFica.BackColor = 16777215
    'Forms![Pyrl - Check Detail]!txtLastFica.SpecialEffect = 2
    'Forms![Pyrl - Check Detail]!txtLastFica.TabStop = True
    'Forms![Pyrl - Check Detail]!txtLastFit.Locked = False
    'Forms![Pyrl - Check Detail]!txtLastFit.BackColor = 16777215
    'Forms![Pyrl - Check Detail]!txtLastFit.SpecialEffect = 2
    'Forms![Pyrl - Check Detail]!txtLastFit.TabStop = True

End Sub

Private Sub cmdPrintChecks_Click()
'Dim d As Database
Dim rsRegister As ADODB.Recordset
Dim rsGLTrans As ADODB.Recordset
'Dim rsGLTransDetail As ADODB.Recordset
Dim rsGLDetailWork As ADODB.Recordset, rsSumGLDetail As ADODB.Recordset, rsEmployees As ADODB.Recordset, rsReg As ADODB.Recordset
Dim SQLstatement As String

If CheckEmpty(True) = True Then Exit Sub
ShowStatus True
'On Error GoTo CmdPrintChecks_Click_Error

'Dim mypath$                     'Reset global database variables
'Dim rsReset As ADODB.Recordset
'Set db = CurrentDb
'Set rsReset = db.OpenRecordset("Last Company")
'If Not rsReset.RecordCount = 0 Then
'  rsReset.MoveFirst
'  mypath$ = rsReset("Last Company")
'  Set db2 = DBEngine.Workspaces(0).OpenDatabase(mypath$)
'End If


Dim AdditionsToGross
ShowStatus True

Set rsPyrlSetup = New ADODB.Recordset
rsPyrlSetup.Open "SELECT [FICA EXP ACCT],[FICA PAY ACCT],[FUTA EXP ACCT],[FUTA PAY ACCT]," & _
"[FIT PAY ACCT],[SUI EXP ACCT],[SUI PAY ACCT],[OFFICE EXP ACCT],[SALES EXP ACCT]," & _
"[WHSE EXP ACCT],[PROD EXP ACCT] FROM [Pyrl - Setup]", db, adOpenKeyset, adLockOptimistic, adCmdText

Dim RecCount As Integer

For RecCount = 0 To rsPyrlSetup.Fields.count - 1
    If IsNull(rsPyrlSetup(RecCount)) Then
        MsgBox "There is an Empty field in Payroll Setup." & vbCr & _
               "Please complete Payroll Setup before continue", vbInformation, "Information"
               frm_SYS_Setup_Payroll.Show
               ShowStatus False
               Exit Sub
    End If
Next

Set rsReg = New ADODB.Recordset
rsReg.Open "Select  * From [Pyrl - Register] Where [Printed] = 0 Order by [ID]", db, adOpenKeyset, adLockOptimistic, adCmdText
    If rsReg.RecordCount = 0 Then
       MsgBox "There are no checks to print", vbInformation, "Error"
       ShowStatus False
       Exit Sub
    End If

'If txtFieldsDate(0) = "" Then
'    ShowStatus False
'    MsgBox "Please enter a CheckDate." & Chr(10) & "This field prints on checks and sets the posting period"
'    Exit Sub
'End If
    

  'Verify period can be posted to; Send TranDate; Return PeriodToPost and PeriodClosed
  Dim PeriodToPost%
  Dim PeriodClosed%
  Call VerifyPeriod(txtFieldsDate(0), PeriodToPost%, PeriodClosed%)
  If PeriodClosed% = True Then
    MsgBox "Unable to post transaction to a closed period.", , "Post Payment Error"
    ShowStatus False
    Exit Sub
  End If


Dim NewDocNumber
Dim ID
Dim Success%
Dim DebitAmount@
Dim CreditAmount@
Dim Account$
Dim TransNumber
Dim TranDate
Dim GLAmount
Dim CRIT As String, sql As String
'Set d = CurrentDb
'Assign Check Numbers
  
  Dim BankID$
  Dim CheckNo$
  Dim FirstCheckNo&
  Dim NumChecks%
  Dim rsHeader As ADODB.Recordset
  Dim ThisCheck&

'Get next check number
  If optCheck(0).Value = True Then
        NextCheck$ = CheckNumberCHQ("READ", db, txtfields.Text)
  Else
GetCheckRange:
        ShowStatus False
        NextCheck$ = InputBox("Enter check number. If you want to end the process, Please hit the cancel button.", , NextCheck$)
        FirstCheckNo& = Val(NextCheck$)
        
        If NextCheck$ = "" Then Exit Sub
        ShowStatus True
        If Not IsNumeric(NextCheck$) Then
            MsgBox "Please enter a valid check number! Only numbers accepted.", , "Error"
            GoTo GetCheckRange
        End If
        Dim CheckStatus As String
        If CheckNumberCHQ("CHECK", db, txtfields.Text, NextCheck$) = "Found" Then
            MsgBox "Please enter a valid check number! It's already been used.", , "Error"
            GoTo GetCheckRange
        End If
  End If
'Exit Sub
    
'Set rsGLTrans = New ADODB.Recordset
'rsGLTrans.Open "[GL Transaction]", db, adOpenKeyset, adLockOptimistic, adCmdTable

'Set rsGLTransDetail = New ADODB.Recordset
'rsGLTransDetail.Open "[GL Transaction Detail]", db, adOpenKeyset, adLockOptimistic, adCmdTable

Set rsGLDetailWork = New ADODB.Recordset
rsGLDetailWork.Open "[Pyrl - GL Trans Detail Work]", db, adOpenKeyset, adLockOptimistic, adCmdTable

Set rsEmployees = New ADODB.Recordset
rsEmployees.Open "[Pyrl - Employee Data]", db, adOpenKeyset, adLockOptimistic, adCmdTable

gLinesPosted% = 0
GLAmount = 0
  
'  Dim rsBank As ADODB.Recordset            '---------------
'  Set rsBank = New ADODB.Recordset
'  rsBank.Open "[Bank Accounts]", db, adOpenKeyset, adLockOptimistic, adCmdTable
  
  'rsBank.Index = "PrimaryKey"
'  rsBank.MoveFirst
'  rsBank.Find "[BANK ACCT ID]='" & txtFields & "'"
'  If rsBank.EOF Then
'    MsgBox "Bank account is not valid!", , "Error"
'    ShowStatus False
'    Exit Sub
'  End If
'GetCheckRange:
    'IIf(optCheck(0).Value = 1, "Computer", "Manual")
'    NextCheck$ = rsBank![BANK ACCT Next Check No]
'    If optCheck(1).Value = True Then
'        ShowStatus False
'        NextCheck$ = InputBox("Enter check number.", , NextCheck$)
'        FirstCheckNo& = Val(NextCheck$)
'    End If
'    If NextCheck$ = "" Then Exit Sub
'    ShowStatus true
'    If Not IsNumeric(NextCheck$) Then
'        MsgBox "Please enter a valid check number!", , "Error"
'        GoTo GetCheckRange
'    End If
 
 'Make sure checks in the desired range are not used and write to register
  'Set rsHeader = New ADODB.Recordset      '------1---AP Payment Header
  'rsHeader.Open "SELECT [AP PAY Check No] FROM [AP Payment Header] WHERE [AP PAY Bank Account]='" & txtFields & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
  'rsHeader.Index = "BankKey"
  'rsReg.MoveLast
  rsReg.MoveFirst
  ThisCheck& = Val(NextCheck$)
  'MsgBox rsReg.RecordCount
  Do While Not rsReg.EOF
'GetCheckRange:
        CheckNo$ = Trim(CStr(ThisCheck&))
        'BankID$ = CStr(txtFields)
        'CheckNumberCHQ
'        If CheckDocument("SELECT [AP PAY Check No] FROM [AP Payment Header] WHERE [AP PAY Bank Account]='" & txtFields & "' AND [AP PAY Check No]='" & CheckNo$ & "'", True) = False Then
        'rsHeader.MoveFirst
        'rsHeader.Find "[AP PAY Check No]=" & CheckNo$ & "'"
        'If rsHeader.EOF Then
            'rsReg'.Edit
            rsReg![CheckNumber] = CheckNo$
            rsReg!CHECKTYPE = IIf(optCheck(0).Value = 1, "Computer", "Manual")
            rsReg!CHECKDATE = txtFieldsDate(0)
            rsReg.Update
'            rsBank![BANK ACCT Next Check No] = CheckNo$ + 1
'            rsBank.Update
'        Else
'            ThisCheck& = ThisCheck& + 1
'            GoTo GetCheckRange
'        End If
'        ThisCheck& = ThisCheck& + 1
        rsReg.MoveNext
  Loop
            MsgBox "Check Number is " & CheckNo$, vbInformation, "Information"
 
  rsReg.MoveFirst
    With rsEmployees
    Do Until rsReg.EOF
        CRIT = "[EMP ID]=" & "'" & rsReg![EMP ID] & "'"    'Post last Pay Date to employee Table
        .MoveFirst
        .Find CRIT
        If Not .EOF Then
            ''.Edit
                !YTDGROSS = !YTDGROSS + rsReg!GROSS   'post YTD Amounts to Employee Table
                !YTDREGHOURS = !YTDREGHOURS + rsReg!REGHOURS
                !YTDOTHOURS = !YTDOTHOURS + rsReg!OTHOURS
            If !FITYN = -1 Then
                !YTDFIT = !YTDFIT + rsReg!FIT
            End If
                !YTDSTATETAX = !YTDSTATETAX + rsReg!STATETAX
                !YTDLOCAL = !YTDLOCAL + rsReg!LOCAL
            If !FICAYN = -1 Then
                ![YTDFICA] = (![YTDFICA] + rsReg!FICA)
            End If
            .Update
        End If
        rsReg.MoveNext
    Loop
    End With
 
db.BeginTrans  '/////////////////////////////////////////////////
'On Error GoTo PostError

 'Reset next check number-
    'rsBank'.Edit
    'rsBank("BANK ACCT Next Check No") = ThisCheck&
    'rsBank.Update

'Write to AP Payment Header
rsReg.MoveFirst
  Do While Not rsReg.EOF
    'rsHeader.AddNew          '----------------
      
    '  rsHeader("AP PAY Type") = "Payroll"
    '  rsHeader("AP PAY Check No") = rsReg![CHECKNUMBER]
    '  rsHeader("AP PAY Vendor No") = rsReg![EMP ID]
    '  rsHeader("AP PAY Transaction Date") = rsReg!CHECKDATE
    '  rsHeader("AP PAY Amount") = rsReg!NETPAY
    '  rsHeader("AP PAY UnApplied Amount") = 0
    '  rsHeader("AP PAY Bank Account") = txtFields.Text
    '  rsHeader("AP PAY Status") = "Posted"
    '  rsHeader("AP PAY Void") = False
    '  rsHeader("AP PAY Notes") = "Paid through payroll."
    '  rsHeader("AP PAY Credit Amount") = 0
    '  rsHeader("AP PAY Class") = 0
    '  rsHeader("AP PAY Cleared") = False
    '  rsHeader("AP PAY Posted YN") = True
    '  rsHeader("AP PAY Recurring YN") = False
      
    'rsHeader.Update
    
  SQLstatement = "INSERT INTO [AP Payment Header]"
  SQLstatement = SQLstatement & " ([AP PAY Type],[AP PAY Check No],[AP PAY Vendor No],"
  SQLstatement = SQLstatement & "[AP PAY Transaction Date],[AP PAY Amount],[AP PAY UnApplied Amount],"
  SQLstatement = SQLstatement & "[AP PAY Bank Account],[AP PAY Status],[AP PAY Void],"
  SQLstatement = SQLstatement & "[AP PAY Notes],[AP PAY Credit Amount],[AP PAY Class],"
  SQLstatement = SQLstatement & "[AP PAY Cleared],[AP PAY Posted YN],[AP PAY Recurring YN])"
  
  SQLstatement = SQLstatement & " VALUES ('Payroll','" & rsReg![CheckNumber] & "','" & rsReg![EMP ID] & "',"
  SQLstatement = SQLstatement & "#" & rsReg!CHECKDATE & "#," & rsReg!NETPAY & ",0,"
  SQLstatement = SQLstatement & "'" & txtfields.Text & "','Posted',FALSE,"
  SQLstatement = SQLstatement & "'Paid through payroll.',0,0,"
  SQLstatement = SQLstatement & "FALSE,TRUE,FALSE)"
  'Debug.Print SQLstatement
  
  db.Execute SQLstatement
  'db.RollbackTrans
  'Exit Sub
 rsReg.MoveNext
 Loop
    
rsReg.MoveFirst
    Do Until rsReg.EOF
        CRIT = "[EMP ID]=" & "'" & rsReg![EMP ID] & "'"    'Post last Pay Date to employee Table
        With rsEmployees
        .MoveFirst
        .Find CRIT
        If Not .EOF Then
            '.Edit
          If IsNull(![LASTDATE]) Then
            ![LASTDATE] = rsReg![Date]
          ElseIf ![LASTDATE] < rsReg![Date] Then
            ![LASTDATE] = rsReg![Date]
          End If
        .Update
        End If
        End With
        rsReg.MoveNext
    Loop
    
    'Get next GL Doc#
    Dim rsDocNumber As ADODB.Recordset
    Set rsDocNumber = New ADODB.Recordset
    rsDocNumber.Open "SELECT [GL TRANS Document #] FROM [GL Transaction] WHERE [GL TRANS Type]='Payroll' ORDER BY [GL TRANS Document #] DESC", db, adOpenKeyset, adLockOptimistic, adCmdText
    With rsDocNumber
    If .RecordCount = 0 Then
     NewDocNumber = "Pyrl-1"
    Else
        .MoveFirst
        NewDocNumber = "Pyrl-" & Right(![GL TRANS Document #], Len(![GL TRANS Document #]) - 5)
    End If
    .Close
    End With
    Set rsDocNumber = Nothing
    
    'write to gl table and get transnumber / auto number
    Set rsRegister = New ADODB.Recordset
    rsRegister.Open "[Pyrl - Register Query]", db, adOpenKeyset, adLockOptimistic, adCmdTable 'Has a record for each pyrl item
    'MsgBox "Using Query, Check this out = " & rsRegister.RecordCount
    If rsRegister.RecordCount = 0 Then
       MsgBox ("There are no checks to print")
        ShowStatus False
        Exit Sub
    End If
     rsRegister.MoveFirst
     'rsGLTrans.AddNew
     '   rsGLTrans![GL TRANS Description] = "Payroll"
     '   rsGLTrans![GL TRANS Date] = txtFieldsDate(0)
     '   rsGLTrans![GL TRANS Type] = "Payroll"
     '   rsGLTrans![GL TRANS Posted YN] = -1
     '   rsGLTrans![GL TRANS Description] = "Payroll Entry"
     '   rsGLTrans![GL TRANS Reference] = "Payroll Entry"
     '   rsGLTrans![GL TRANS Document #] = NewDocNumber
     'rsGLTrans.Update
     '   MsgBox "Used adUseServer"

  SQLstatement = "INSERT INTO [GL Transaction]"
  SQLstatement = SQLstatement & " ([GL TRANS Description],[GL TRANS Date],[GL TRANS Type],[GL TRANS Posted YN],"
  SQLstatement = SQLstatement & "[GL TRANS Reference],[GL TRANS Document #])"
  SQLstatement = SQLstatement & "VALUES ('Payroll Entry',#" & txtFieldsDate(0) & "#,"
  SQLstatement = SQLstatement & "'Payroll',-1,'Payroll Entry',"
  SQLstatement = SQLstatement & "'" & NewDocNumber & "')"
  'Debug.Print SQLstatement
  
  db.Execute SQLstatement
    Set rsGLTrans = New ADODB.Recordset
    rsGLTrans.Open "SELECT [GL TRANS Number] FROM [GL Transaction] WHERE [GL TRANS Document #] ='" & NewDocNumber & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
        TransNumber = rsGLTrans![GL TRANS Number]
    rsGLTrans.Close
    Set rsGLTrans = Nothing
        
'Write the gl WORK TABLE  detail
'******************************************************
    Do Until rsRegister.EOF
        ID = rsRegister![Pyrl - Register.ID]
        
        gLinesPosted% = gLinesPosted% + 1
        With rsGLDetailWork
            
            'Debit FICA Employer Expense Account
            '.AddNew
            '![GL TRANSD Number] = TransNumber
            '    If IsNull(rsPyrlSetup![FICA EXP ACCT]) Then
            '        MsgBox ("Enter a FICA Expense Account in Payroll Setup!")
            '        GoTo PostError
            '    End If
            '![GL TRANSD Account] = rsPyrlSetup![FICA EXP ACCT]
            '![GL TRANSD Debit Amount] = rsRegister!FICAER
            '![GL TRANSD Credit Amount] = 0
            '![GL TRANSD Project] = 0
            '![BankAcctNumber] = txtFields
            '![CHECKNUMBER] = rsRegister!CHECKNUMBER
            ' .Update
                SQLstatement = "INSERT INTO [Pyrl - GL Trans Detail Work]"
                SQLstatement = SQLstatement & " ([GL TRANSD Number],[GL TRANSD Account],[GL TRANSD Debit Amount],"
                SQLstatement = SQLstatement & "[GL TRANSD Credit Amount],[GL TRANSD Project],[BankAcctNumber],[CHECKNUMBER])"
                
                SQLstatement = SQLstatement & "VALUES (" & TransNumber & ",'" & rsPyrlSetup![FICA EXP ACCT] & "',"
                SQLstatement = SQLstatement & rsRegister!FICAER & ",0,0,'" & txtfields.Text & "',"
                SQLstatement = SQLstatement & rsRegister!CheckNumber & ")"
                'Debug.Print SQLstatement
                
                db.Execute SQLstatement
            
            Success% = PostCOA(rsPyrlSetup![FICA EXP ACCT], txtFieldsDate(0), rsRegister!FICAER, 0)
            If Success% = False Then GoTo PostError
        
            'Debit FUTA Employer Expense Account
            '.AddNew
            '![GL TRANSD Number] = TransNumber
            '    If IsNull(rsPyrlSetup![FUTA EXP ACCT]) Then
            '        MsgBox ("Enter a FUTA Expense Account in Payroll Setup!")
            '        GoTo PostError
            '    End If
            '![GL TRANSD Account] = rsPyrlSetup![FUTA EXP ACCT]
            '![GL TRANSD Debit Amount] = rsRegister!FUTA
            '![GL TRANSD Credit Amount] = 0
            '![GL TRANSD Project] = 0
            '![BankAcctNumber] = txtFields
            '![CHECKNUMBER] = rsRegister!CHECKNUMBER
            ' .Update
                SQLstatement = "INSERT INTO [Pyrl - GL Trans Detail Work]"
                SQLstatement = SQLstatement & " ([GL TRANSD Number],[GL TRANSD Account],[GL TRANSD Debit Amount],"
                SQLstatement = SQLstatement & "[GL TRANSD Credit Amount],[GL TRANSD Project],[BankAcctNumber],[CHECKNUMBER])"
                
                SQLstatement = SQLstatement & "VALUES (" & TransNumber & ",'" & rsPyrlSetup![FUTA EXP ACCT] & "',"
                SQLstatement = SQLstatement & rsRegister!FUTA & ",0,0,'" & txtfields & "',"
                SQLstatement = SQLstatement & rsRegister!CheckNumber & ")"
                'Debug.Print SQLstatement
                
                db.Execute SQLstatement
            Success% = PostCOA(rsPyrlSetup![FUTA EXP ACCT], txtFieldsDate(0), NZ(rsRegister!FUTA, 0), 0)
            If Success% = False Then GoTo PostError
           
           'Debit SUI Employer Expense Account
            '.AddNew
            '![GL TRANSD Number] = TransNumber
            '    If IsNull(rsPyrlSetup![SUI EXP ACCT]) Then
            '        MsgBox ("Enter a SUI Expense Account in Payroll Setup!")
                    
            '        GoTo PostError
            '    End If
            '![GL TRANSD Account] = rsPyrlSetup![SUI EXP ACCT]
            '![GL TRANSD Debit Amount] = rsRegister!SUI
            '![GL TRANSD Credit Amount] = 0
            '![GL TRANSD Project] = 0
            '![BankAcctNumber] = txtFields
            '![CHECKNUMBER] = rsRegister!CHECKNUMBER
            ' .Update
                SQLstatement = "INSERT INTO [Pyrl - GL Trans Detail Work]"
                SQLstatement = SQLstatement & " ([GL TRANSD Number],[GL TRANSD Account],[GL TRANSD Debit Amount],"
                SQLstatement = SQLstatement & "[GL TRANSD Credit Amount],[GL TRANSD Project],[BankAcctNumber],[CHECKNUMBER])"
                
                SQLstatement = SQLstatement & "VALUES (" & TransNumber & ",'" & rsPyrlSetup![SUI EXP ACCT] & "',"
                SQLstatement = SQLstatement & rsRegister!SUI & ",0,0,'" & txtfields & "',"
                SQLstatement = SQLstatement & rsRegister!CheckNumber & ")"
                'Debug.Print SQLstatement
                
                db.Execute SQLstatement
            
            Success% = PostCOA(rsPyrlSetup![SUI EXP ACCT], txtFieldsDate(0), NZ(rsRegister!SUI, 0), 0)
            If Success% = False Then GoTo PostError
           'Credit FICA Employee and Employer Liab Account
            '.AddNew
            '![GL TRANSD Number] = TransNumber
            '   If IsNull(rsPyrlSetup![FICA PAY ACCT]) Then
            '        MsgBox ("Enter a FICA Payable Account in Payroll Setup!")
            '        GoTo PostError
            '   End If
            '![GL TRANSD Account] = rsPyrlSetup![FICA PAY ACCT]
            '![GL TRANSD Debit Amount] = 0
            '![GL TRANSD Credit Amount] = (rsRegister![FICAER] + rsRegister![FICA])
            '![GL TRANSD Project] = 0
            '![BankAcctNumber] = txtFields
            '![CHECKNUMBER] = rsRegister!CHECKNUMBER
            ' .Update
                
                SQLstatement = "INSERT INTO [Pyrl - GL Trans Detail Work]"
                SQLstatement = SQLstatement & " ([GL TRANSD Number],[GL TRANSD Account],[GL TRANSD Debit Amount],"
                SQLstatement = SQLstatement & "[GL TRANSD Credit Amount],[GL TRANSD Project],[BankAcctNumber],[CHECKNUMBER])"
                
                SQLstatement = SQLstatement & "VALUES (" & TransNumber & ",'" & rsPyrlSetup![FICA PAY ACCT] & "',"
                SQLstatement = SQLstatement & "0," & (rsRegister![FICAER] + rsRegister![FICA]) & ",0,'" & txtfields & "',"
                SQLstatement = SQLstatement & rsRegister!CheckNumber & ")"
                'Debug.Print SQLstatement
                
                db.Execute SQLstatement
            
            Success% = PostCOA(rsPyrlSetup![FICA PAY ACCT], txtFieldsDate(0), 0, rsRegister!FICAER + rsRegister!FICA)
            If Success% = False Then GoTo PostError
           
          'Credit Employee FIT payable acct
           '.AddNew
           '  ![GL TRANSD Number] = TransNumber
           '         If IsNull(rsPyrlSetup![FIT PAY ACCT]) Then
           '             MsgBox ("Enter a FIT Payable Account in Payroll Setup!")
           '             GoTo PostError
           '         End If
           ' ![GL TRANSD Account] = rsPyrlSetup![FIT PAY ACCT]
           ' ![GL TRANSD Debit Amount] = 0
           ' ![GL TRANSD Credit Amount] = rsRegister!FIT
           ' ![GL TRANSD Project] = 0
           ' ![BankAcctNumber] = txtFields
           ' ![CHECKNUMBER] = rsRegister!CHECKNUMBER
           '.Update
           
                SQLstatement = "INSERT INTO [Pyrl - GL Trans Detail Work]"
                SQLstatement = SQLstatement & " ([GL TRANSD Number],[GL TRANSD Account],[GL TRANSD Debit Amount],"
                SQLstatement = SQLstatement & "[GL TRANSD Credit Amount],[GL TRANSD Project],[BankAcctNumber],[CHECKNUMBER])"
                
                SQLstatement = SQLstatement & "VALUES (" & TransNumber & ",'" & rsPyrlSetup![FIT PAY ACCT] & "',"
                SQLstatement = SQLstatement & "0," & rsRegister!FIT & ",0,'" & txtfields & "',"
                SQLstatement = SQLstatement & rsRegister!CheckNumber & ")"
                'Debug.Print SQLstatement
                
                db.Execute SQLstatement
                
        Success% = PostCOA(NZ(rsPyrlSetup![FIT PAY ACCT], ""), txtFieldsDate(0), 0, rsRegister!FIT)
        If Success% = False Then GoTo PostError
        
          'Credit Employer Futa payable acct
          ' .AddNew
          '   ![GL TRANSD Number] = TransNumber
          '          If IsNull(rsPyrlSetup![FUTA PAY ACCT]) Then
          '              MsgBox ("Enter a FUTA Payable Account in Payroll Setup!")
          '              GoTo PostError
          '          End If
          '  ![GL TRANSD Account] = rsPyrlSetup![FUTA PAY ACCT]
          '  ![GL TRANSD Debit Amount] = 0
          '  ![GL TRANSD Credit Amount] = rsRegister!FUTA
          '  ![GL TRANSD Project] = 0
          '  ![BankAcctNumber] = txtFields
          '  ![CHECKNUMBER] = rsRegister!CHECKNUMBER
          ' .Update
                SQLstatement = "INSERT INTO [Pyrl - GL Trans Detail Work]"
                SQLstatement = SQLstatement & " ([GL TRANSD Number],[GL TRANSD Account],[GL TRANSD Debit Amount],"
                SQLstatement = SQLstatement & "[GL TRANSD Credit Amount],[GL TRANSD Project],[BankAcctNumber],[CHECKNUMBER])"
                
                SQLstatement = SQLstatement & "VALUES (" & TransNumber & ",'" & rsPyrlSetup![FUTA PAY ACCT] & "',"
                SQLstatement = SQLstatement & "0," & rsRegister!FUTA & ",0,'" & txtfields & "',"
                SQLstatement = SQLstatement & rsRegister!CheckNumber & ")"
                'Debug.Print SQLstatement
                
                db.Execute SQLstatement
        Success% = PostCOA(NZ(rsPyrlSetup![FUTA PAY ACCT], ""), txtFieldsDate(0), 0, NZ(rsRegister!FUTA, 0))
        If Success% = False Then GoTo PostError
         
        'Credit Employer SUI payable acct
           '.AddNew
           '  ![GL TRANSD Number] = TransNumber
           '         If IsNull(rsPyrlSetup![SUI PAY ACCT]) Then
           '             MsgBox ("Enter a SUI Payable Account in Payroll Setup!")
           '             GoTo PostError
           '         End If
           ' ![GL TRANSD Account] = rsPyrlSetup![SUI PAY ACCT]
           ' ![GL TRANSD Debit Amount] = 0
           ' ![GL TRANSD Credit Amount] = rsRegister!SUI
           ' ![GL TRANSD Project] = 0
           ' ![BankAcctNumber] = txtFields
           ' ![CHECKNUMBER] = rsRegister!CHECKNUMBER
           '.Update
                SQLstatement = "INSERT INTO [Pyrl - GL Trans Detail Work]"
                SQLstatement = SQLstatement & " ([GL TRANSD Number],[GL TRANSD Account],[GL TRANSD Debit Amount],"
                SQLstatement = SQLstatement & "[GL TRANSD Credit Amount],[GL TRANSD Project],[BankAcctNumber],[CHECKNUMBER])"
                
                SQLstatement = SQLstatement & "VALUES (" & TransNumber & ",'" & rsPyrlSetup![SUI PAY ACCT] & "',"
                SQLstatement = SQLstatement & "0," & rsRegister!SUI & ",0,'" & txtfields & "',"
                SQLstatement = SQLstatement & rsRegister!CheckNumber & ")"
                'Debug.Print SQLstatement
                
                db.Execute SQLstatement
        Success% = PostCOA(rsPyrlSetup![SUI PAY ACCT], txtFieldsDate(0), 0, NZ(rsRegister!SUI, 0))
        If Success% = False Then GoTo PostError
           
        'Credit Cash acct
           '.AddNew
           '  ![GL TRANSD Number] = TransNumber
           '             If IsNull(txtFields) Then
           '                 MsgBox ("Enter a Cash Account Number!")
           '            GoTo PostError
           '         End If
           '  ![GL TRANSD Account] = txtFields
           '  ![GL TRANSD Debit Amount] = 0
           '  ![GL TRANSD Credit Amount] = rsRegister!NETPAY
           '  ![GL TRANSD Project] = 0
           '![BankAcctNumber] = txtFields
           ' ![CHECKNUMBER] = rsRegister!CHECKNUMBER
           '.Update
                SQLstatement = "INSERT INTO [Pyrl - GL Trans Detail Work]"
                SQLstatement = SQLstatement & " ([GL TRANSD Number],[GL TRANSD Account],[GL TRANSD Debit Amount],"
                SQLstatement = SQLstatement & "[GL TRANSD Credit Amount],[GL TRANSD Project],[BankAcctNumber],[CHECKNUMBER])"
                
                SQLstatement = SQLstatement & "VALUES (" & TransNumber & ",'" & txtfields.Text & "',"
                SQLstatement = SQLstatement & "0," & rsRegister!NETPAY & ",0,'" & txtfields & "',"
                SQLstatement = SQLstatement & rsRegister!CheckNumber & ")"
                'Debug.Print SQLstatement
                
                db.Execute SQLstatement
                
        Success% = PostCOA(txtfields, txtFieldsDate(0), 0, rsRegister!NETPAY)
        If Success% = False Then GoTo PostError

        'Employee Payroll Items
         AdditionsToGross = 0
         Do While rsRegister![Pyrl - Register.ID] = ID
            If Not NZ(rsRegister!PyrlItemID, "") = "" Then        'skip query records with no pyrl items
            
                If NZ(rsRegister!Account, "") = "" Then         'pyrl item exists but has no gl account
                    MsgBox ("Payroll item" & "'  " & rsRegister!Description & "'  " & "needs an account number")
                    GoTo PostError
                End If
              
                '.AddNew
                '        ![GL TRANSD Number] = TransNumber
                '        ![GL TRANSD Account] = rsRegister!Account
                '        ![BankAcctNumber] = txtFields
                '        ![CHECKNUMBER] = rsRegister!CHECKNUMBER
                '        ![GL TRANSD Project] = 0
                        SQLstatement = "INSERT INTO [Pyrl - GL Trans Detail Work]"
                        SQLstatement = SQLstatement & " ([GL TRANSD Number],[GL TRANSD Account],"
                        SQLstatement = SQLstatement & "[BankAcctNumber],[CHECKNUMBER],[GL TRANSD Project],"
                        
                
                Select Case rsRegister!Type     'Additions
                    Case "Addition"
                        '![GL TRANSD Debit Amount] = NZ(rsRegister!TotalAmount, 0)
                        '![GL TRANSD Credit Amount] = 0
                        '.Update
                        SQLstatement = SQLstatement & "[GL TRANSD Debit Amount],[GL TRANSD Credit Amount])"
                        
                        SQLstatement = SQLstatement & "VALUES (" & TransNumber & ",'" & rsRegister!Account & "',"
                        SQLstatement = SQLstatement & "'" & txtfields & "'," & rsRegister!CheckNumber & ",0,"
                        SQLstatement = SQLstatement & NZ(rsRegister!TotalAmount, 0) & ",0)"
                        'Debug.Print SQLstatement
                        db.Execute SQLstatement
                        
                         If Not rsRegister!Basis = "Net" Then
                            AdditionsToGross = AdditionsToGross + NZ(rsRegister!TotalAmount, 0) 'Accum Gross/AGI Additions to reduce Gross Wages exp posted
                         End If
                        Success% = PostCOA(rsRegister!Account, txtFieldsDate(0), NZ(rsRegister!TotalAmount, 0), 0)
                            
                           'Post Employer Debit if any
                            If rsRegister!EmployerYN = -1 Then
                                If NZ(rsRegister!Account2, "") = "" Then         'Employer Item is selected but has no gl account
                                    MsgBox ("Payroll item" & "'  " & rsRegister!Description & "'  " & "needs an employer account number")
                                    GoTo PostError
                                End If
                                '.AddNew
                                '![GL TRANSD Number] = TransNumber
                                '![GL TRANSD Account] = rsRegister!Account2
                                '![BankAcctNumber] = txtFields
                                '![CHECKNUMBER] = rsRegister!CHECKNUMBER
                                '![GL TRANSD Project] = 0
                                '![GL TRANSD Debit Amount] = NZ(rsRegister!TotalAmountEmployer, 0)
                                '![GL TRANSD Credit Amount] = 0
                                '.Update
                                
                                SQLstatement = "INSERT INTO [Pyrl - GL Trans Detail Work]"
                                SQLstatement = SQLstatement & " ([GL TRANSD Number],[GL TRANSD Account],"
                                SQLstatement = SQLstatement & "[BankAcctNumber],[CHECKNUMBER],[GL TRANSD Project],"
                                SQLstatement = SQLstatement & "[GL TRANSD Debit Amount],[GL TRANSD Credit Amount])"
                                
                                SQLstatement = SQLstatement & "VALUES (" & TransNumber & ",'" & rsRegister!Account2 & "',"
                                SQLstatement = SQLstatement & "'" & txtfields & "'," & rsRegister!CheckNumber & ",0,"
                                SQLstatement = SQLstatement & NZ(rsRegister!TotalAmountEmployer, 0) & ",0)"
                                'Debug.Print SQLstatement
                                db.Execute SQLstatement
                                
                                Success% = PostCOA(rsRegister!Account2, txtFieldsDate(0), NZ(rsRegister!TotalAmountEmployer, 0), 0)
                            End If
                
                             'Post Employer Credit if any
                            If rsRegister!EmployerYN = -1 Then
                                If NZ(rsRegister!Account3, "") = "" Then         'Employer Item is selected but has no gl account
                                    MsgBox ("Payroll item" & "'  " & rsRegister!Description & "'  " & "needs an employer account number")
                                    GoTo PostError
                                End If
                                '.AddNew
                                '![GL TRANSD Number] = TransNumber
                                '![GL TRANSD Account] = rsRegister!Account3
                                '![BankAcctNumber] = txtFields
                                '![CHECKNUMBER] = rsRegister!CHECKNUMBER
                                '![GL TRANSD Project] = 0
                                '![GL TRANSD Debit Amount] = 0    'Deductions
                                '![GL TRANSD Credit Amount] = NZ(rsRegister!TotalAmountEmployer, 0)
                                '.Update
                                
                                SQLstatement = "INSERT INTO [Pyrl - GL Trans Detail Work]"
                                SQLstatement = SQLstatement & " ([GL TRANSD Number],[GL TRANSD Account],"
                                SQLstatement = SQLstatement & "[BankAcctNumber],[CHECKNUMBER],[GL TRANSD Project],"
                                SQLstatement = SQLstatement & "[GL TRANSD Debit Amount],[GL TRANSD Credit Amount])"
                                
                                SQLstatement = SQLstatement & "VALUES (" & TransNumber & ",'" & rsRegister!Account3 & "',"
                                SQLstatement = SQLstatement & "'" & txtfields & "'," & rsRegister!CheckNumber & ",0,"
                                SQLstatement = SQLstatement & "0," & NZ(rsRegister!TotalAmountEmployer, 0) & ")"
                                'Debug.Print SQLstatement
                                db.Execute SQLstatement
                                
                                Success% = PostCOA(rsRegister!Account3, txtFieldsDate(0), 0, NZ(rsRegister!TotalAmountEmployer, 0))
                            End If
                    
                    Case "Deduction"
                        '![GL TRANSD Debit Amount] = 0    'Deductions
                        '![GL TRANSD Credit Amount] = NZ(rsRegister!TotalAmount, 0)
                        '.Update
                        
                        SQLstatement = SQLstatement & "[GL TRANSD Debit Amount],[GL TRANSD Credit Amount])"
                        
                        SQLstatement = SQLstatement & "VALUES (" & TransNumber & ",'" & rsRegister!Account & "',"
                        SQLstatement = SQLstatement & "'" & txtfields & "'," & rsRegister!CheckNumber & ",0,"
                        SQLstatement = SQLstatement & "0," & NZ(rsRegister!TotalAmount, 0) & ")"
                        'Debug.Print SQLstatement
                        db.Execute SQLstatement
                        
                        Success% = PostCOA(rsRegister!Account, txtFieldsDate(0), 0, NZ(rsRegister!TotalAmount, 0))
                
                            'Post Employer Debit if any
                            If rsRegister!EmployerYN = -1 Then
                                If NZ(rsRegister!Account2, "") = "" Then         'Employer Item is selected but has no gl account
                                    MsgBox ("Payroll item" & "'  " & rsRegister!Description & "'  " & "needs an employer account number")
                                    GoTo PostError
                                End If
                                '.AddNew
                                '![GL TRANSD Number] = TransNumber
                                '![GL TRANSD Account] = rsRegister!Account2
                                '![BankAcctNumber] = txtFields
                                '![CHECKNUMBER] = rsRegister!CHECKNUMBER
                                '![GL TRANSD Project] = 0
                                '![GL TRANSD Debit Amount] = NZ(rsRegister!TotalAmountEmployer, 0)
                                '![GL TRANSD Credit Amount] = 0
                                '.Update
                                
                                SQLstatement = "INSERT INTO [Pyrl - GL Trans Detail Work]"
                                SQLstatement = SQLstatement & " ([GL TRANSD Number],[GL TRANSD Account],"
                                SQLstatement = SQLstatement & "[BankAcctNumber],[CHECKNUMBER],[GL TRANSD Project],"
                                SQLstatement = SQLstatement & "[GL TRANSD Debit Amount],[GL TRANSD Credit Amount])"
                                
                                SQLstatement = SQLstatement & "VALUES (" & TransNumber & ",'" & rsRegister!Account2 & "',"
                                SQLstatement = SQLstatement & "'" & txtfields & "'," & rsRegister!CheckNumber & ",0,"
                                SQLstatement = SQLstatement & NZ(rsRegister!TotalAmountEmployer, 0) & ",0)"
                                'Debug.Print SQLstatement
                                db.Execute SQLstatement
                                
                                Success% = PostCOA(rsRegister!Account2, txtFieldsDate(0), NZ(rsRegister!TotalAmountEmployer, 0), 0)
                            End If
                           
                           'Post Employer Credit if any
                            If rsRegister!EmployerYN = -1 Then
                                If NZ(rsRegister!Account3, "") = "" Then         'Employer Item is selected but has no gl account
                                    MsgBox ("Payroll item" & "'  " & rsRegister!Description & "'  " & "needs an employer account number")
                                    GoTo PostError
                                End If
                                '.AddNew
                                '![GL TRANSD Number] = TransNumber
                                '![GL TRANSD Account] = rsRegister!Account3
                                '![BankAcctNumber] = txtFields
                                '![CHECKNUMBER] = rsRegister!CHECKNUMBER
                                '![GL TRANSD Project] = 0
                                '![GL TRANSD Debit Amount] = 0    'Deductions
                                '![GL TRANSD Credit Amount] = NZ(rsRegister!TotalAmountEmployer, 0)
                                '.Update
                                
                                SQLstatement = "INSERT INTO [Pyrl - GL Trans Detail Work]"
                                SQLstatement = SQLstatement & " ([GL TRANSD Number],[GL TRANSD Account],"
                                SQLstatement = SQLstatement & "[BankAcctNumber],[CHECKNUMBER],[GL TRANSD Project],"
                                SQLstatement = SQLstatement & "[GL TRANSD Debit Amount],[GL TRANSD Credit Amount])"
                                
                                SQLstatement = SQLstatement & "VALUES (" & TransNumber & ",'" & rsRegister!Account3 & "',"
                                SQLstatement = SQLstatement & "'" & txtfields & "'," & rsRegister!CheckNumber & ",0,"
                                SQLstatement = SQLstatement & "0," & NZ(rsRegister!TotalAmountEmployer, 0) & ")"
                                'Debug.Print SQLstatement
                                db.Execute SQLstatement
                                
                                Success% = PostCOA(rsRegister!Account3, txtFieldsDate(0), 0, NZ(rsRegister!TotalAmountEmployer, 0))
                            End If
                
                Case "State Tax"
                        '![GL TRANSD Debit Amount] = 0    'Deductions
                        '![GL TRANSD Credit Amount] = NZ(rsRegister!TotalAmount, 0)
                        '.Update
                        
                        SQLstatement = SQLstatement & "[GL TRANSD Debit Amount],[GL TRANSD Credit Amount])"
                        
                        SQLstatement = SQLstatement & "VALUES (" & TransNumber & ",'" & rsRegister!Account & "',"
                        SQLstatement = SQLstatement & "'" & txtfields & "'," & rsRegister!CheckNumber & ",0,"
                        SQLstatement = SQLstatement & "0," & NZ(rsRegister!TotalAmount, 0) & ")"
                        'Debug.Print SQLstatement
                        db.Execute SQLstatement
                        
                        Success% = PostCOA(rsRegister!Account, txtFieldsDate(0), 0, NZ(rsRegister!TotalAmount, 0))
                
                Case "Local Tax"
                        '![GL TRANSD Debit Amount] = 0    'Deductions
                        '![GL TRANSD Credit Amount] = NZ(rsRegister!TotalAmount, 0)
                        '.Update
                        
                        SQLstatement = SQLstatement & "[GL TRANSD Debit Amount],[GL TRANSD Credit Amount])"
                        
                        SQLstatement = SQLstatement & "VALUES (" & TransNumber & ",'" & rsRegister!Account & "',"
                        SQLstatement = SQLstatement & "'" & txtfields & "'," & rsRegister!CheckNumber & ",0,"
                        SQLstatement = SQLstatement & "0," & NZ(rsRegister!TotalAmount, 0) & ")"
                        'Debug.Print SQLstatement
                        db.Execute SQLstatement
                        
                        Success% = PostCOA(rsRegister!Account, txtFieldsDate(0), 0, NZ(rsRegister!TotalAmount, 0))
                 End Select
                If Success% = False Then GoTo PostError
            
            End If
           
            rsRegister.MoveNext
                If rsRegister.EOF Then
                    Exit Do
                End If
          Loop      'Loop payroll items
        
        rsRegister.MovePrevious
        'Debit wages expense acct
          '.AddNew
          '  ![GL TRANSD Number] = TransNumber
            Select Case rsRegister!DEPARTMENT
                Case "Office"
                    If IsNull(rsPyrlSetup![OFFICE EXP ACCT]) Then
                        MsgBox ("Enter GL posting accounts in payroll setup!")
                    GoTo PostError
                    Else
                    Account$ = rsPyrlSetup![OFFICE EXP ACCT]
                    End If
                
                Case "Sales"
                    If IsNull(rsPyrlSetup![SALES EXP ACCT]) Then
                        MsgBox ("Enter GL posting accounts in payroll setup!")
                    GoTo PostError
                    Else
                    Account$ = rsPyrlSetup![SALES EXP ACCT]
                    End If
                
                Case "Warehouse"
                    If IsNull(rsPyrlSetup![WHSE EXP ACCT]) Then
                        MsgBox ("Enter GL posting accounts in payroll setup!")
                    GoTo PostError
                    Else
                    Account$ = rsPyrlSetup![WHSE EXP ACCT]
                    End If
                
                Case "Production"
                    If IsNull(rsPyrlSetup![PROD EXP ACCT]) Then
                        MsgBox ("Enter GL posting accounts in payroll setup!")
                    GoTo PostError
                    Else
                    Account$ = rsPyrlSetup![PROD EXP ACCT]
                    End If
            End Select
           ' ![GL TRANSD Account] = NZ(Account$, "")
           ' ![GL TRANSD Debit Amount] = (rsRegister!GROSS - AdditionsToGross)
           ' ![GL TRANSD Credit Amount] = 0
           ' ![GL TRANSD Project] = 0
           ' ![BankAcctNumber] = txtFields
           ' ![CHECKNUMBER] = rsRegister!CHECKNUMBER
           ' .Update
            SQLstatement = "INSERT INTO [Pyrl - GL Trans Detail Work]"
            SQLstatement = SQLstatement & " ([GL TRANSD Number],[GL TRANSD Account],"
            SQLstatement = SQLstatement & "[GL TRANSD Debit Amount],[GL TRANSD Credit Amount],"
            SQLstatement = SQLstatement & "[GL TRANSD Project],[BankAcctNumber],[CHECKNUMBER])"
                                
            SQLstatement = SQLstatement & "VALUES (" & TransNumber & ",'" & NZ(Account$, "") & "',"
            SQLstatement = SQLstatement & (rsRegister!GROSS - AdditionsToGross) & ",0,"
            SQLstatement = SQLstatement & "0,'" & txtfields & "'," & rsRegister!CheckNumber & ")"
            'Debug.Print SQLstatement
            db.Execute SQLstatement
'-----------------------------------------------------------------------------------
            Success% = PostCOA(NZ(Account$, ""), txtFieldsDate(0), (rsRegister!GROSS - rsRegister![Pyrl - Register.ADDITIONS]), 0)
            If Success% = False Then GoTo PostError
            rsRegister.MoveNext 'return to last record for this employee
        End With
            'MsgBox rsRegister.AbsolutePosition
            'MsgBox rsRegister.RecordCount
        Loop                            'Loop Register query recordset
   
    'write to GL Detail Table from work table
    Dim cmdGLDetail As Command
    'Dim rsSumGLDetail As ADODB.Recordset
    Dim ParamGLDetail As Parameter
    
    Set cmdGLDetail = New Command
    cmdGLDetail.ActiveConnection = db
    
    cmdGLDetail.CommandText = "[Pyrl - SumGLDetailWork]"
    cmdGLDetail.CommandType = adCmdStoredProc
    
    Set ParamGLDetail = cmdGLDetail.CreateParameter("CurrentDoc", adInteger, adParamInput) 'Screen.ActiveForm.[AP PAY Check No]       'set query criteria for current work table records
    
    ParamGLDetail.Value = TransNumber
    
    cmdGLDetail.Parameters.Append ParamGLDetail
    
    Set rsSumGLDetail = cmdGLDetail.Execute

    'Dim QdfGLDetail As QueryDef
    'Set QdfGLDetail = d.QueryDefs("Pyrl - SumGLDetailWork")
    'QdfGLDetail.Parameters![currentDoc] = TransNumber       'set query criteria for current work table records
    'Set rsSumGLDetail = QdfGLDetail.OpenRecordset
    
    'With rsGLTransDetail
    rsSumGLDetail.MoveFirst
    Do Until rsSumGLDetail.EOF
        
        SQLstatement = "INSERT INTO [GL Transaction Detail]"
        SQLstatement = SQLstatement & " ([GL TRANSD Number],[GL TRANSD Account],"
        SQLstatement = SQLstatement & "[GL TRANSD Debit Amount],[GL TRANSD Credit Amount],[GL TRANSD Project])"
        SQLstatement = SQLstatement & "VALUES (" & TransNumber & ",'" & rsSumGLDetail!Account & "',"
        SQLstatement = SQLstatement & rsSumGLDetail!Debit & "," & rsSumGLDetail!Credit & ",0)"
        'Debug.Print SQLstatement
        
        db.Execute SQLstatement
        '.AddNew
        '![GL TRANSD Number] = TransNumber
        '![GL TRANSD Account] = rsSumGLDetail!Account
        '![GL TRANSD Debit Amount] = rsSumGLDetail!Debit
        '![GL TRANSD Credit Amount] = rsSumGLDetail!Credit
        '![GL TRANSD Project] = 0
        '.Update
        GLAmount = GLAmount + rsSumGLDetail!Debit
        rsSumGLDetail.MoveNext
            If rsSumGLDetail.EOF Then   'ni apa cite ni nak buat berapa kali
            Exit Do
            End If
        'MsgBox rsSumGLDetail.AbsolutePosition
        'MsgBox rsSumGLDetail.RecordCount
    Loop
    'End With

db.Execute ("UPDATE [GL Transaction] SET [GL Transaction].[GL TRANS Amount] = " & GLAmount & " WHERE ((([GL Transaction].[GL TRANS Number])= " & TransNumber & "))")

'Mark Commissionable invoices used as paid in AR Sales table
db.Execute "UPDATE [Pyrl - Register Detail Commissions] INNER JOIN [AR Sales] ON [Pyrl - Register Detail Commissions].ExtDocNo = [AR Sales].[AR SALE Ext Document #] SET [AR Sales].[AR SALE Commission Paid] = -1"

'Me.Requery
'Me.Refresh  'refresh bank acct total
lblCheckPrints = "$0.00"

 'Print Checks
    
    If optCheck(0).Value = True Then 'Computer checks
        MsgBox ("Make sure checks are ascending order in the printer and press OK.")
        'DoCmd.OpenReport "Pyrl - Checks", acViewNormal
        
        If MsgBox("Did the checks print correctly?" & Chr$(10) & "Selecting 'No' will void ALL checks for this printing.", vbYesNo) = vbNo Then
            GoTo PrintError
        End If
    Else
        If Check1.Value = 1 Then
            'DoCmd.OpenReport "Pyrl - Manual Checks Posted", acNormal
        End If
    
    End If

db.Execute "Update [Pyrl - Register] Set [Printed] = -1 Where [Printed] = 0"
'Set d = Nothing
    
db.CommitTrans  'It's all over baby ///////////////////////////////

If optCheck(0).Value = True Then
    MsgBox (" Payroll Transactions Posted.")
Else
    MsgBox "Payroll Transactions Posted;" & Chr$(10) & "You may proceed with writing manual checks."
End If

ShowStatus False
'rsGLTransDetail.Close
rsPyrlSetup.Close
rsGLDetailWork.Close
rsSumGLDetail.Close
rsEmployees.Close
rsReg.Close
rsRegister.Close
Call RedoPurchaseNumbers(db)

Exit Sub

PostError:
 db.RollbackTrans
 'return the check number by minus one
 CheckNumberCHQ "BACK", db, txtfields.Text
 
 db.Execute "Update [Pyrl - Register] Set [CHECKNUMBER] = 0 Where [Printed] = 0"
 db.Execute "Update [Pyrl - Register] Set [CHECKDATE] = Null Where [Printed] = 0"
 db.Execute "Update [Pyrl - Register] Set [CHECKTYPE] = Null Where [Printed] = 0"
rsReg.MoveFirst
 'Forms![Pyrl - Pay Employees].[Pyrl - Employee Data subform1].Form.Requery
    With rsEmployees
    Do Until rsReg.EOF
        CRIT = "[EMP ID]=" & "'" & rsReg![EMP ID] & "'"    'Post last Pay Date to employee Table
        .MoveFirst
        .Find CRIT
        If Not .EOF Then
            '.Edit
            !YTDGROSS = !YTDGROSS - rsReg!GROSS   'post YTD Amounts to Employee Table
            !YTDREGHOURS = !YTDREGHOURS - rsReg!REGHOURS
            !YTDOTHOURS = !YTDOTHOURS - rsReg!OTHOURS
                    If !FITYN = -1 Then
            !YTDFIT = !YTDFIT - rsReg!FIT
                End If
            !YTDSTATETAX = !YTDFIT - rsReg!STATETAX
            !YTDLOCAL = !YTDFIT - rsReg!LOCAL
                If !FICAYN = -1 Then
            ![YTDFICA] = (![YTDFICA] - rsReg!FICA)
                End If
            .Update
        End If
        rsReg.MoveNext
    Loop
    End With

ShowStatus False
 MsgBox "An error occurred; Transaction not posted!", , "Error"
 GoTo CmdPrintChecks_Click_Error

PrintError:
 
  db.RollbackTrans
  CheckNumberCHQ "BACK", db, txtfields.Text
  
    db.Execute "Update [Pyrl - Register] Set [CHECKNUMBER] = 0 Where [Printed] = 0"
    db.Execute "Update [Pyrl - Register] Set [CHECKDATE] = null Where [Printed] = 0"
    db.Execute "Update [Pyrl - Register] Set [CHECKTYPE] = null Where [Printed] = 0"
  
  rsReg.MoveFirst
    With rsEmployees
    Do Until rsReg.EOF
        CRIT = "[EMP ID]=" & "'" & rsReg![EMP ID] & "'"    'Post last Pay Date to employee Table
        .MoveFirst
        .Find CRIT
        If Not .EOF Then
            '.Edit
                !YTDGROSS = !YTDGROSS - rsReg!GROSS   'post YTD Amounts to Employee Table
                !YTDREGHOURS = !YTDREGHOURS - rsReg!REGHOURS
                !YTDOTHOURS = !YTDOTHOURS - rsReg!OTHOURS
            If !FITYN = -1 Then
                !YTDFIT = !YTDFIT - rsReg!FIT
            End If
                !YTDSTATETAX = !YTDFIT - rsReg!STATETAX
                !YTDLOCAL = !YTDFIT - rsReg!LOCAL
            If !FICAYN = -1 Then
                ![YTDFICA] = (![YTDFICA] - rsReg!FICA)
            End If
            .Update
        End If
        rsReg.MoveNext
    Loop
    End With
  
  'Reset next check number-
    
'    Set rsBank = New ADODB.Recordset
'    rsBank.Open "[Bank Accounts]", db, adOpenKeyset, adLockOptimistic, adCmdTable
    'rsBank.Index = "PrimaryKey"
'    rsBank.MoveFirst
'    rsBank.Find "[BANK ACCT ID]='" & txtFields & "'"
    'rsBank .Edit
'    rsBank("BANK ACCT Next Check No") = ThisCheck&
'    rsBank.Update
  
  'Mark printed checks as void in AP Payment Header
    'Set d = CurrentDb
    'Set rsRegister = New ADODB.Recordset
    'rsRegister.Open "Select * From [Pyrl - Register] where [Printed] = 0 Order by [ID]", db, adOpenKeyset, adLockOptimistic, adCmdText
        'Set rsHeader = New ADODB.Recordset        '--------------
        'rsHeader.Open "[AP Payment Header]", db, adOpenKeyset, adLockOptimistic, adCmdTable
        'rsHeader.Index = "BankKey"
    'rsRegister.MoveFirst
    rsReg.MoveFirst
    Do While Not rsReg.EOF
        CheckNo$ = Trim(CStr(FirstCheckNo&))
        'rsHeader.AddNew
        '  rsHeader("AP PAY Type") = "Payroll"
        '  rsHeader("AP PAY Check No") = CheckNo$
        '  rsHeader("AP PAY Vendor No") = rsRegister![EMP ID]
        '  rsHeader("AP PAY Transaction Date") = rsRegister!Date
        '  rsHeader("AP PAY Amount") = rsRegister!GROSS
        '  rsHeader("AP PAY UnApplied Amount") = 0
        '  rsHeader("AP PAY Bank Account") = txtFields
        '  rsHeader("AP PAY Status") = "" '--------------------
        '  rsHeader("AP PAY Void") = True
        '  rsHeader("AP PAY Notes") = "Paid through payroll."
        '  rsHeader("AP PAY Credit Amount") = 0
        '  rsHeader("AP PAY Class") = 0
        '  rsHeader("AP PAY Cleared") = False
        '  rsHeader("AP PAY Posted YN") = False
        '  rsHeader("AP PAY Recurring YN") = False
        '  rsHeader("AP PAY Status") = "Open" '------------------
        'rsHeader.Update
    
      SQLstatement = "INSERT INTO [AP Payment Header]"
      SQLstatement = SQLstatement & " ([AP PAY Type],[AP PAY Check No],[AP PAY Vendor No],"
      SQLstatement = SQLstatement & "[AP PAY Transaction Date],[AP PAY Amount],[AP PAY UnApplied Amount],"
      SQLstatement = SQLstatement & "[AP PAY Bank Account],[AP PAY Status],[AP PAY Void],"
      SQLstatement = SQLstatement & "[AP PAY Notes],[AP PAY Credit Amount],[AP PAY Class],"
      SQLstatement = SQLstatement & "[AP PAY Cleared],[AP PAY Posted YN],[AP PAY Recurring YN])"
      
      SQLstatement = SQLstatement & " VALUES ('Payroll','" & CheckNo$ & "','" & rsReg![EMP ID] & "',"
      SQLstatement = SQLstatement & "#" & rsReg!Date & "#," & rsReg!GROSS & ",0,"
      SQLstatement = SQLstatement & "'" & txtfields.Text & "','Open',True,"
      SQLstatement = SQLstatement & "'Paid through payroll.',0,0,"
      SQLstatement = SQLstatement & "False,False,False)"
      'Debug.Print SQLstatement
      
      db.Execute SQLstatement
    
      rsRegister.MoveNext
      FirstCheckNo& = FirstCheckNo& + 1
 Loop
 
 MsgBox "Your Checks have been Voided"
 
 'rsGLTransDetail.Close
 rsRegister.Close
 rsPyrlSetup.Close
 rsGLDetailWork.Close
 rsEmployees.Close
 rsReg.Close
 'rsBank.Close
 'Set d = Nothing
 'Me.Refresh
 'Forms![Pyrl - Pay Employees].[Pyrl - Employee Data subform1].Form.Requery
 ShowStatus False
Exit Sub

CmdPrintChecks_Click_Error:
ShowStatus False
Call ErrorLog("GL Entry", "cmdPost_Click", Now, Err.Number, Err.Description, True, db)
'Resume Next
End Sub
Private Sub cmdPyrlItems_Click()

If CheckEmpty = True Then Exit Sub

'On Error GoTo Err_cmdPyrlItems_Click

    'Dim stDocName As String
    'Dim stLinkCriteria As String

    'stDocName = "Pyrl - Payroll Items"
    frPayrollItems.ZOrder 0
    frPayrollItems.Visible = True
    Me.Caption = "Payroll Items"
    Form_Resize
    frEmployee.Enabled = False
    frAccount.Enabled = False
    chkPyrllItem.Enabled = False
    LoadPayrollDB
    
    'DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_cmdPyrlItems_Click:
    Exit Sub

Err_cmdPyrlItems_Click:
    MsgBox Err.Description
    Resume Exit_cmdPyrlItems_Click

End Sub

Private Sub PayrollItemDataBase()
    'open database
End Sub

Private Sub CmdSelect_Click()
'If TempEMPID = "" Then
'    MsgBox "Please make your selection", vbInformation, "Information"
'    Exit Sub
'End If
grdDataGrid.SetFocus
'Dim dbs As Database,
'Dim rst As ADODB.Recordset
'Dim strSQL$, Typeselected$, Frequency$

'On Error GoTo cmdSelect_Click_Error
    
'     Set dbs = CurrentDb
    
    'Open Employee Data recordset
'    Frequency = Combo1.Text
'    Typeselected = Combo2.Text
    
'        If Typeselected = "All" Then
'            strSQL$ = "SELECT * FROM [Pyrl - Employee Data]"
'        Else
'            strSQL$ = "SELECT * FROM [Pyrl - Employee Data] where [PAYTYPE]= '"
'            strSQL$ = strSQL$ & Typeselected
'            strSQL$ = strSQL$ & "'"
'        End If
         
'        If Frequency = "All" Then
'            strSQL$ = strSQL$
'        Else
'            If Typeselected = "All" Then
'            strSQL$ = "SELECT * FROM [Pyrl - Employee Data] where [PAYFREQUENCY] = '"
'            Else
'            strSQL$ = strSQL$ & " and [PAYFREQUENCY] = '"
'            End If
'        strSQL$ = strSQL$ & Frequency
'        strSQL$ = strSQL$ & "'"
'        End If
    
'    Set rst = New ADODB.Recordset
'    rst.Open strSQL$, db, adOpenKeyset, adLockOptimistic, adCmdText
'    If rst.RecordCount = 0 Then
'        MsgBox "There were no employees for this selection criteria", vbInformation, "Error"
'        Exit Sub
'    End If
  LoadSelect = True
  With ADOprimaryrs
     .MoveFirst
        Do Until .EOF
        'rst '.Edit
         If ![PAY] = 0 Then
            TempEMPID = grdDataGrid.Columns(1)
            ![PAY] = -1
            .Update
            'SendKeys ("{ENTER}")
            CalcPayroll
         End If
         .MoveNext
        Loop
            Set grdDataGrid.DataSource = Nothing
               ADOprimaryrs.Requery
            Set grdDataGrid.DataSource = ADOprimaryrs
            
        .MoveFirst
        Do Until .EOF
            If grdDataGrid.Columns(3) = 0 Then
                TempEMPID = grdDataGrid.Columns(1)
                MsgBox "Check Amount for " & TempEMPID & " is " & grdDataGrid.Columns(3) & vbCr & "There is no payment to be made.", vbInformation, "Information"
                ![PAY] = 0
                .Update
            End If
            .MoveNext
        Loop
        
  End With
  LoadSelect = False
    'rst.Close
    'Set dbs = Nothing
  
  'Me.[Pyrl - Employee Data subform1].Form.Requery
'  Call CalcPayroll
  'Me.[Pyrl - Employee Data subform1].Form.Requery
  'Me![Pyrl - Employee Data subform1].Form.Refresh
  
  
'  Exit Sub

'cmdSelect_Click_Error:
'  Call ErrorLog("Pay Employees", "CmdSelect_Click", Now,  Err.number, Err.description, True,db)
'  Resume Next

End Sub

Private Function CbValidate(cbName As ComboBox, cbText As String) As Boolean
Dim i As Integer
    For i = 0 To cbName.ListCount - 1
        If cbText = cbName.List(i) Then
            CbValidate = True
            Exit Function
        End If
    Next
    cbName.Text = cbName.List(0)
    CbValidate = False
End Function

Private Sub Combo1_KeyPress(KeyAscii As Integer)
Dim keyResponse As Boolean
    keyResponse = CtrlValidate(KeyAscii, "")
    If keyResponse = True Then
    Else
       KeyAscii = 0
    End If
End Sub

Private Sub Combo1_LostFocus()
    If CbValidate(Combo1, Combo1.Text) = False Then
       MsgBox "There is no such selection", vbInformation, "Information"
    End If
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
Dim keyResponse As Boolean
    keyResponse = CtrlValidate(KeyAscii, "")
    If keyResponse = True Then
    Else
       KeyAscii = 0
    End If
End Sub

Private Sub Combo2_LostFocus()
    If CbValidate(Combo2, Combo2.Text) = False Then
       MsgBox "There is no such selection", vbInformation, "Information"
    End If
End Sub

Private Sub Form_Resize()
If frCommision.Visible = True Then
    Me.Width = 12015
Else
    Me.Width = 9825
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo FormErr
  ShowStatus True
  If ADOprimaryrs.State = 0 Then GoTo JumpSkip
      If ADOprimaryrs.RecordCount > 0 Then
        If ADOprimaryrs.EditMode <> 0 Then
          ADOprimaryrs.CancelUpdate
        End If
      End If
      ADOprimaryrs.Close
      Set ADOprimaryrs = Nothing
JumpSkip:
      db.Close
      Set db = Nothing
  ShowStatus False
  Exit Sub
FormErr:
  MsgBox Err.Description
  ShowStatus False
End Sub

Private Sub grdDataGrid_AfterColEdit(ByVal ColIndex As Integer)
    If grdDataGrid.Row = -1 Or grdDataGrid.Columns(0) = "" Then Exit Sub
      SendKeys ("{ENTER}")
  If grdDataGrid.Row > 0 Then
      SendKeys ("{up}")
      SendKeys ("{down}")
  ElseIf grdDataGrid.Row = 0 Then
      SendKeys ("{down}")
      SendKeys ("{up}")
  End If
End Sub

Private Sub grdDataGrid_ButtonClick(ByVal ColIndex As Integer)
On Error GoTo Error_ButtClick
ShowStatus True
If grdDataGrid.Row < 0 Then Exit Sub
If grdDataGrid.Row > -1 And LoadSelect = False Then
    lblName = LookRecord("[EMP Name]", "[EMP Employees]", db, "[EMP ID] = '" & grdDataGrid.Columns(1) & "'")
    TempEMPID = grdDataGrid.Columns(1)
End If
Select Case ColIndex
Case 0
         SendKeys ("{ENTER}")
   If grdDataGrid.Columns(0).Text = "No" Then
      grdDataGrid.Columns(0).Text = "Yes"
   Else
      grdDataGrid.Columns(0).Text = "No"
   End If
         SendKeys ("{ENTER}")
         SendKeys ("{down}")
         SendKeys ("{up}")
Case 6
    Menu_Calendar.WhoCallMe True, 1525
    'Menu_Calendar.Show vbModal
End Select
grdDataGrid_AfterColEdit ColIndex
ADOprimaryrs.Update
If grdDataGrid.Columns(0).Text = "No" Then
    ShowStatus False
    Exit Sub
Else
      Dim mvBookMark
      LoadSelect = True
      With ADOprimaryrs
        CalcPayroll
        'Set grdDataGrid.DataSource = Nothing
        If Not (.BOF Or .EOF) Then
          mvBookMark = .Bookmark
        End If
        .Requery
        If mvBookMark > 0 Then
          .Bookmark = mvBookMark
        End If
        'Set grdDataGrid.DataSource = ADOprimaryrs
        If grdDataGrid.Columns(3) = 0 Then
            MsgBox "Check Amount for " & TempEMPID & " is " & grdDataGrid.Columns(3) & vbCr & "There is no payment to be made.", vbInformation, "Information"
            ![PAY] = 0
            .Update
        End If
      End With
      LoadSelect = False
End If
ShowStatus False
Exit Sub
Error_ButtClick:
    MsgBox "Please click the Table box before clicking the button"
End Sub

Private Sub grdDataGrid_Error(ByVal DataError As Integer, Response As Integer)
    If DataGridKnownError(DataError) Then
        Response = 0
    End If
End Sub

Private Sub grdDataGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If grdDataGrid.Row < 0 Then Exit Sub
If grdDataGrid.Row > -1 And LoadSelect = False Then
    lblName = LookRecord("[EMP Name]", "[EMP Employees]", db, "[EMP ID] = '" & grdDataGrid.Columns(1) & "'")
    TempEMPID = grdDataGrid.Columns(1)
End If
Select Case grdDataGrid.col
  Case 2
     grdDataGrid.AllowUpdate = True
  Case Else
     grdDataGrid.AllowUpdate = False
  End Select
Exit Sub
Damn_Attempt:
     grdDataGrid.AllowUpdate = False
     grdDataGrid.col = 0
exit_sub:

End Sub

Private Sub optCheck_Click(Index As Integer)
Select Case Index
Case 0
    Check1.Enabled = False
Case 1
    Check1.Enabled = True
End Select
End Sub

Private Sub txtFieldsDate_Change(Index As Integer)
Select Case Index
Case 1, 2
If txtFieldsDate(1) <> "" And txtFieldsDate(2) <> "" Then
    If txtFieldsDate(1) > txtFieldsDate(2) Then
        MsgBox "Period Beginning Date cannot be later than Pay Period Ending Date"
        txtFieldsDate(2) = ""
        Exit Sub
    End If
End If
End Select
End Sub

Private Sub OpenRecordsets(Optional Track As Boolean, Optional ProcessName As String, Optional EMPIDdata As String)
Dim strSQL As String

'On Error GoTo OpenRecordsets_Error
  
   'Open Employee Data, select marked records as a group, or pay employee from check detail
    If Me.Caption = "Payroll - Check Detail" Then
        strSQL = "SELECT * FROM [Pyrl - Employees] WHERE [EMP ID]= '" & TempEMPID & "'"
    Else
        If LoadSelect = False Then
            strSQL = "SELECT * FROM [Pyrl - Employees] WHERE [PAY] = -1 and [EMP ID] not like 'TEMPLATE'"
        Else
            strSQL = "SELECT * FROM [Pyrl - Employees] WHERE [PAY] = -1 and [EMP ID]='" & TempEMPID & "'"
        End If
        'strSQL = "SELECT * FROM [Pyrl - Employees] WHERE [PAY] = -1 "
    End If
    
Set Rst = New ADODB.Recordset
    Rst.Open strSQL, db, adOpenKeyset, adLockOptimistic, adCmdText
        

        If Rst.RecordCount = 0 Then
            If Me.Caption = "Sales Commission" Then
                Track = False
            Else
                MsgBox ("There are no Employees Selected For Payment")
            End If
            ShowStatus False
            Exit Sub
        End If
        Rst.MoveFirst
  
  'Open Payroll Register recordset
Set rstPyrl = New ADODB.Recordset
    rstPyrl.Open "[Pyrl - Register]", db, adOpenKeyset, adLockOptimistic, adCmdTable
   
   'Open Payroll Setup recordset
Set rstSETUP = New ADODB.Recordset
    rstSETUP.Open "[Pyrl - Setup]", db, adOpenKeyset, adLockOptimistic, adCmdTable
    If rstSETUP.RecordCount > 0 Then
        rstSETUP.MoveFirst
    Else
        MsgBox ("Please enter Payroll Setup before calculating payroll checks.")
        ShowStatus False
        End
    End If
       
    'Open Commissions recordset
    If LookRecord("[PAYALLCOMMISSIONS]", "[Pyrl - Setup]", db) = -1 Then
        'Mark Invoices and Returns for all commissionable invoices
        db.Execute "Update [AR SALES] SET [AR SALE Select to Pay]= 0 WHERE [AR SALE Commission Paid] = 0"
        'DoCmd.SetWarnings False
'-------------------------------------look this        'DoCmd.OpenQuery ("Pyrl - AutoPayCommission")
        strSQL = "UPDATE [AR SALES] SET [AR SALE Select To Pay] = -1 "
        strSQL = strSQL & "WHERE [AR SALE Commission Paid])=0 AND [AR SALE Date]<=" & txtFieldsDate(2)
        db.Execute strSQL
'-------------------------------------
        'DoCmd.SetWarnings True
    End If
        
    'Calculates based on manual or auto  selected invoices
Set rsCommissions = New ADODB.Recordset
    rsCommissions.Open "[Pyrl - Commission]", db, adOpenKeyset, adLockOptimistic, adCmdTable
    If rsCommissions.RecordCount > 0 Then
        rsCommissions.MoveFirst
    End If

Exit Sub
OpenRecordsets_Error:
Call ErrorLog("Payroll Module", "OpenRecordsets", Now, Err.Number, Err.Description, True, db)

End Sub

'
Private Sub PostToRegister()
  
  Dim strSQL As String
  Dim rstRD As ADODB.Recordset, rstInvoices As ADODB.Recordset, rstRDC As ADODB.Recordset
  Dim i, Number
        
  Call OpenRecordsets
        
   'Open Payroll Register Detail recordset
    Set rstRD = New ADODB.Recordset
    rstRD.Open "[Pyrl - Register Detail]", db, adOpenKeyset, adLockOptimistic, adCmdTable
        
        
'Post NetCheck Amounts
'rst.MoveLast
Rst.MoveFirst

Dim K As Integer

For K = 0 To Rst.RecordCount - 1
        'rst.Edit
        
        Call Formulas
        
        Rst![NETCHECK] = Net
        Rst![PAY] = 0   'Clear select to pay
        Rst.Update
    
   If Net >= 0 Then
    
    'Post to Payroll Register
    'With rstPyrl
    '.AddNew
    '![EMP ID] = rst![EMP ID]
    '![Name] = rst![EMP Name]
    '!Date = txtFieldsDate(2)
    '!StartDate = txtFieldsDate(1)
    '!GROSS = GROSS
    '!PRETAXDED = PRETAXDED
    '!YTDGROSS = rst!YTDGROSS + GROSS
    '!AGI = AGI
    '!FIT = FIT
    '!FICA = FicaEE
    '!FICAER = FICAER
    '!FUTA = FUTA
    '!SUI = SUI
    '!STATETAX = STATETAX
    '!LOCAL = LocalTax
    '!NETPAY = Net
    '!ADDITIONS = ADDITIONS
    '!DEDUCTIONS = DEDUCTIONS
    '!Commission = Commission
    '!OTHOURS = OT
    '!REGHOURS = Regular
    '!Amount = rst!Amount    'WhichType=1
    
  Dim SQLstatement As String
  
  SQLstatement = "INSERT INTO [Pyrl - Register]"
  SQLstatement = SQLstatement & " ([EMP ID],[Name],[Date],[StartDate],"
  SQLstatement = SQLstatement & "[GROSS],[PRETAXDED],[YTDGROSS],[AGI],[FIT],[FICA],"
  SQLstatement = SQLstatement & "[FICAER],[FUTA],[SUI],[STATETAX],[LOCAL],[NETPAY],"
  SQLstatement = SQLstatement & "[ADDITIONS],[DEDUCTIONS],[Commission],"
  SQLstatement = SQLstatement & "[OTHOURS],[REGHOURS],[Amount],[OTRATE],[HOURLYRATE])"
  
  SQLstatement = SQLstatement & "VALUES ('" & TempEMPID & "','" & lblName & "',#" & txtFieldsDate(2).Text & "#,#" & txtFieldsDate(1).Text & "#,"
  SQLstatement = SQLstatement & GROSS & "," & PRETAXDED & "," & Rst!YTDGROSS + GROSS & "," & AGI & "," & FIT & "," & FicaEE & ","
  SQLstatement = SQLstatement & FICAER & "," & FUTA & "," & SUI & "," & STATETAX & "," & LocalTax & "," & Net & ","
  SQLstatement = SQLstatement & ADDITIONS & "," & DEDUCTIONS & "," & Commission & ","
  SQLstatement = SQLstatement & OT & "," & Regular & "," & Rst!Amount & ","
    
    If Me.Caption = "Payroll - Check Detail" Then
'    If isloaded("Pyrl - Check Detail") Then
        If ckCheckDetail.Value = 1 Then
            '!OTRATE = rst!OTRATE
            '!HOURLYRATE = rst!HOURLYRATE
            SQLstatement = SQLstatement & Rst!OTRATE & "," & Rst!HOURLYRATE & ")"
        Else                            'Blank Check
            '!OTRATE = rst!BlankCheckOTRate
            '!HOURLYRATE = rst!BlankCheckHourlyRate
             SQLstatement = SQLstatement & Rst!BlankCheckOTRate & "," & Rst!BlankCheckHourlyRate & ")"
        End If
    Else                                'Standard Check Rates
    '!OTRATE = rst!OTRATE
    '!HOURLYRATE = rst!HOURLYRATE
    SQLstatement = SQLstatement & Rst!OTRATE & "," & Rst!HOURLYRATE & ")"
    End If
    
  'Debug.Print SQLstatement
 
  db.Execute SQLstatement

  Dim Temprs As ADODB.Recordset
  Set Temprs = New ADODB.Recordset
  Temprs.Open "SELECT [ID] FROM [Pyrl - Register] ORDER BY [ID]", db, adOpenKeyset, adLockOptimistic, adCmdText
    Temprs.MoveLast
    Number = Temprs!ID
  Temprs.Close
  Set Temprs = Nothing
    '.Update
    'End With
   
     
       'Post Pyrl Register Detail -Commissions
        
        'Set rstRDC = db.OpenRecordset("Pyrl - Register Detail Commissions")
        
        strSQL = "SELECT [Pyrl - Calc Commissions Unmatched].[EMP ID], [Pyrl - Calc Commissions Unmatched].[AR SALE Ext Document #], [Pyrl - Calc Commissions Unmatched].[AR SALE Document Type], [Pyrl - Calc Commissions Unmatched].Commission"
        strSQL = strSQL & " FROM [Pyrl - Calc Commissions Unmatched] WHERE ((([Pyrl - Calc Commissions Unmatched].[EMP ID])='" & Rst![EMP ID] & "'))" '"
        
        Set rstInvoices = New ADODB.Recordset
        rstInvoices.Open strSQL, db, adOpenKeyset, adLockOptimistic, adCmdText
        
        If rstInvoices.RecordCount > 0 Then
        rstInvoices.MoveFirst
        With rstRDC
        
        For i = 0 To rstInvoices.RecordCount - 1
            SQLstatement = "INSERT INTO [Pyrl - Register Detail Commissions]"
            SQLstatement = SQLstatement & "([ID],[DocType],[ExtDocNo],[Commission])"
            
            SQLstatement = SQLstatement & "VALUES (" & Number & ","
            SQLstatement = SQLstatement & "'" & rstInvoices![AR SALE Document Type] & "',"
            SQLstatement = SQLstatement & "'" & rstInvoices![AR SALE Ext Document #] & "',"
            SQLstatement = SQLstatement & rstInvoices![Commission] & ")"
        
            db.Execute SQLstatement
        '.AddNew
        '!ID = Number
        '!DocType = rstInvoices![AR SALE Document Type]
        '!ExtDocNo = rstInvoices![AR SALE Ext Document #]
        '!Commission = rstInvoices![Commission]
        '.Update
        rstInvoices.MoveNext
        Next i
        End With
        End If
       
       'Post Pyrl Register Detail - Payroll items
       If rstItems.RecordCount > 0 Then
        rstItems.MoveFirst
        
        With rstRD
        For i = 0 To rstItems.RecordCount - 1
        If rstItems!TotalAmount > 0 Then   'Item must have a value to post to detail
            
            '.AddNew
            '!ID = Number      'AutoNumber
            '!PyrlItemID = rstItems!PyrlItemID
            '![EMP ID] = rstItems![EMP ID]
            '!Description = rstItems!Description
            '!Basis = rstItems!Basis
            '!Type = rstItems!Type
            '!YTDMax = rstItems!YTDMax
            '!Minimum = rstItems!Minimum
            '!WageHigh = rstItems!WageHigh
            '!WageLow = rstItems!WageLow
            '!ItemAmount = rstItems!ItemAmount
            '!ItemPercent = rstItems!ItemPercent
            '!TotalAmount = rstItems!TotalAmount
            '!EmployerTotalAmount = rstItems!EmployerTotalAmount
            '!PercentAmount = rstItems!PercentAmount
            '!EmployerPercentAmount = rstItems!EmployerPercentAmount
            '!ApplyItem = rstItems!ApplyItem
            '!Account = rstItems!Account
            '!Account2 = rstItems!Account2
            '!Account3 = rstItems!Account3
            '!EmployerItemPercent = rstItems!EmployerItemPercent
            '!EmployerItemAmount = rstItems!EmployerItemAmount
            '!EmployerYN = rstItems!EmployerYN
            '.Update

            SQLstatement = "INSERT INTO [Pyrl - Register Detail]"
            SQLstatement = SQLstatement & "([ID],[PyrlItemID],[EMP ID],[Description],"
            SQLstatement = SQLstatement & "[Basis],[Type],[YTDMax],[Minimum],"
            SQLstatement = SQLstatement & "[WageHigh],[WageLow],[ItemAmount],[ItemPercent],"
            SQLstatement = SQLstatement & "[TotalAmount],[EmployerTotalAmount],[PercentAmount],[EmployerPercentAmount],"
            If IsNull(rstItems!Account2) Then
                If IsNull(rstItems!Account3) Then
                    SQLstatement = SQLstatement & "[ApplyItem],[Account],"
                Else
                    SQLstatement = SQLstatement & "[ApplyItem],[Account],[Account3],"
                End If
            Else
                SQLstatement = SQLstatement & "[ApplyItem],[Account],[Account2],[Account3],"
            End If
            SQLstatement = SQLstatement & "[EmployerItemPercent],[EmployerItemAmount],[EmployerYN])"
            
            SQLstatement = SQLstatement & "VALUES (" & Number & ","
            SQLstatement = SQLstatement & "'" & rstItems!PyrlItemID & "',"
            SQLstatement = SQLstatement & "'" & rstItems![EMP ID] & "','" & rstItems!Description & "',"
            SQLstatement = SQLstatement & "'" & rstItems!Basis & "','" & rstItems!Type & "'," & rstItems!YTDMax & "," & rstItems!Minimum & ","
            SQLstatement = SQLstatement & rstItems!WageHigh & "," & rstItems!WageLow & "," & rstItems!ItemAmount & "," & rstItems!ItemPercent & ","
            SQLstatement = SQLstatement & rstItems!TotalAmount & "," & rstItems!EmployerTotalAmount & "," & rstItems!PercentAmount & "," & rstItems!EmployerPercentAmount & ","
            If IsNull(rstItems!Account2) Then
                If IsNull(rstItems!Account3) Then
                    SQLstatement = SQLstatement & rstItems!ApplyItem & "," & rstItems!Account & ","
                Else
                    SQLstatement = SQLstatement & rstItems!ApplyItem & "," & rstItems!Account & "," & rstItems!Account3 & ","
                End If
            Else
                SQLstatement = SQLstatement & rstItems!ApplyItem & "," & rstItems!Account & "," & rstItems!Account2 & "," & rstItems!Account3 & ","
            End If
            SQLstatement = SQLstatement & rstItems!EmployerItemPercent & "," & rstItems!EmployerItemAmount & "," & rstItems!EmployerYN & ")"
        
            'Debug.Print SQLstatement
            db.Execute SQLstatement
         End If
         
NxtItem: rstItems.MoveNext
         Next i
         End With
         End If
        
    rstInvoices.Close
    rstItems.Close
    
  End If
 Rst.MoveNext
Next K

'Set d = Nothing

Rst.Close
rstRD.Close
rstSETUP.Close
'rstFIT.Close
rstPyrl.Close
rsCommissions.Close

Set Rst = Nothing
Set rstRD = Nothing
Set rstSETUP = Nothing
Set rstPyrl = Nothing
Set rsCommissions = Nothing

'Forms![Pyrl - Pay Employees]!cmdPreviewPyrl.SetFocus


End Sub

Private Sub CalcPayroll()
'On Error GoTo cmdCalcPyrl_Click_Error
  
   Call OpenRecordsets
    
    Rst.MoveLast
    Numrec = Rst.RecordCount
    Rst.MoveFirst
        
    For K = 0 To (Numrec - 1)
        With Rst
        '.Edit
        
        Call Formulas
        
        ![NETCHECK] = Net
        !LASTGROSS = GROSS
        !LASTAGI = AGI
        !LASTFIT = FIT
        !LASTSTATETAX = STATETAX
        !LASTLOCAL = LocalTax
        !LASTFICA = FicaEE
        !ADDITIONS = ADDITIONS
        !DEDUCTIONS = DEDUCTIONS
        !PRETAXDED = PRETAXDED
        !LASTCOMMISSION = Commission
        .Update
        rstItems.Close
        .MoveNext
        End With
        
        
    Next K
    
Rst.Close
rstSETUP.Close
'rstFIT.Close
rstPyrl.Close
rsCommissions.Close
Set Rst = Nothing
Set rstSETUP = Nothing
Set rstPyrl = Nothing
Set rsCommissions = Nothing

Exit Sub

cmdCalcPyrl_Click_Error:
ShowStatus False
Call ErrorLog("Pay Employees", "cmdCalcPyrl_Click", Now, Err.Number, Err.Description, True, db)

End Sub


Public Sub UpdatePyrlItems(db As ADODB.Connection, EMP_ID As String)

'Delete all Payroll items for selected employee
db.Execute "Delete * From [Pyrl - Empl Payroll Items] Where [EMP ID] = '" & EMP_ID & "'"


'Mark all selected records in work table with employee id
db.Execute "Update [Pyrl - Select Pyrl Items Work] SET [EMP ID] = 'Erase'"
db.Execute "Update [Pyrl - Select Pyrl Items Work] SET [EMP ID] = '" & EMP_ID & "' WHERE [ApplyItem] = -1"


'Copy Work Table to Employee Pyrl items file
db.Execute "Insert Into [Pyrl - Empl Payroll Items]SELECT * FROM [Pyrl - Select Pyrl Items Work] WHERE [ApplyItem] = -1"

End Sub


Private Sub Formulas(Optional FormType As Integer)
'On Error GoTo FormulasError

If Rst!PAYFREQUENCY = "BiWeekly" Then
    OTConversion = NZ(rstSETUP!OTAfter, 0) * 2
Else
    OTConversion = NZ(rstSETUP!OTAfter, 0)
End If

If Rst!LASTHOURS > OTConversion Then       'calc hours and overtime hours
    OT = (Rst!LASTHOURS - OTConversion)
    Regular = OTConversion
Else
    Regular = Rst!LASTHOURS
    OT = 0
End If
  
 'Get Commission amount for current employee
Commission = 0
Criteria = "[EMP ID] = '" & Rst![EMP ID] & "'"

If rsCommissions.RecordCount > 0 Then
    rsCommissions.Find Criteria
    If Not rsCommissions.EOF Then
        Commission = rsCommissions!Commission
    End If
End If
 
 Select Case UCase(Rst!PAYFREQUENCY)
 Case "WEEKLY"      'Weekly
        Periods = 52
 Case "BIWEEKLY"    'BiWeekly
        Periods = 26
 Case "SEMIMONTHLY" 'SemiMonthly
        Periods = 24
 Case "MONTHLY"     'Monthly
        Periods = 12
 Case "YEARLY"      'Yearly
        Periods = 1
End Select
        
    If Me.Caption = "Payroll - Check Detail" Then         'Pyrl - Check Detail
        If ckCheckDetail.Value = 0 Then 'Manual check
            GROSS = (Regular * Rst!BlankCheckHourlyRate) + (OT * Rst!BlankCheckOTRate) + Rst!Amount + Commission
        Else
            GROSS = (Rst!SALARY / Periods) + (Regular * Rst!HOURLYRATE) + (OT * Rst!OTRATE) + Commission
        End If
    Else
            GROSS = (Rst!SALARY / Periods) + (Regular * Rst!HOURLYRATE) + (OT * Rst!OTRATE) + Commission
    End If
    
   
    GROSSDED = 0
    AGIDED = 0
    PRETAXDED = 0
    FIT = 0
    STATETAX = 0
    LocalTax = 0
    FicaEE = 0
    FICAER = 0
    FUTA = 0
    SUI = 0
    ADDITIONS = 0
    NETADDITIONS = 0
    DEDUCTIONS = 0
    AGI = GROSS
    Net = 0

    If Me.Caption = "Payroll - Check Detail" Then       'Open Employee Payroll Items recordset
        strSQL = "Select * from [Pyrl - Select Pyrl Items Work] Where [EMP ID] = '" & Rst![EMP ID] & "'" & " AND [ApplyItem] = -1 ORDER BY [TYPE]"
    Else
        strSQL = "Select * from [Pyrl - Empl Payroll Items] Where [EMP ID] = '" & Rst![EMP ID] & "'" & " AND [ApplyItem] = -1 ORDER BY [TYPE]"
    End If
    Set rstItems = New ADODB.Recordset
    rstItems.Open strSQL, db, adOpenKeyset, adLockOptimistic, adCmdText
    
    '-----------------------------------------------------------------------------------
    '
    Dim cmdMinMax As Command
    Dim rsMinMax As ADODB.Recordset
    Dim ParamMinMax1 As Parameter
    Dim ParamMinMax2 As Parameter
    Dim ParamMinMax3 As Parameter
    
    Set cmdMinMax = New Command
    cmdMinMax.ActiveConnection = db
    
    cmdMinMax.CommandText = "[Pyrl - Item MinMax]"
    cmdMinMax.CommandType = adCmdStoredProc
    
    Set ParamMinMax1 = cmdMinMax.CreateParameter("CurrentEmp", adBSTR, adParamInput) 'Screen.ActiveForm.[AP PAY Check No]       'set query criteria for current work table records
    Set ParamMinMax2 = cmdMinMax.CreateParameter("StartDate", adDate, adParamInput) 'Screen.ActiveForm.[AP PAY Check No]       'set query criteria for current work table records
    Set ParamMinMax3 = cmdMinMax.CreateParameter("EndDate", adDate, adParamInput) 'Screen.ActiveForm.[AP PAY Check No]       'set query criteria for current work table records
    
    ParamMinMax1.Value = Rst![EMP ID]
    ParamMinMax2.Value = LookRecord("[SYS COM Fiscal Start Date]", "[SYS Company]", db)
    ParamMinMax3.Value = LookRecord("[SYS COM Fiscal End Date]", "[SYS Company]", db)
    
    'MsgBox ParamMinMax1.Value
    'MsgBox ParamMinMax2.Value
    'MsgBox ParamMinMax3.Value
    
    cmdMinMax.Parameters.Append ParamMinMax1
    cmdMinMax.Parameters.Append ParamMinMax2
    cmdMinMax.Parameters.Append ParamMinMax3
    
    Set rsMinMax = cmdMinMax.Execute
    'MsgBox rsMinMax.RecordCount
    
    'Dim Qdf As QueryDef
    'Dim rsMinMax As ADODB.Recordset
    'Set Qdf = db.QueryDefs("Pyrl - Item MinMax")
    'Qdf.Parameters![CurrentEmp] = rst![EMP ID]       'set Employee criteria for MinMax query/recordset
    'Set rsMinMax = Qdf.OpenRecordset
    
    
If rstItems.RecordCount > 0 Then
            rstItems.MoveFirst
            Do Until rstItems.EOF
                If rstItems![Basis] = "Gross" Then  'Items with basis of gross only
                    
                        If TestHighLow(Rst!YTDGROSS, rstItems!WageLow, rstItems!WageHigh) = False Then 'Test High/Low
                            'rstItems.Edit
                             rstItems!TotalAmount = 0
                             rstItems!PercentAmount = 0
                             rstItems!EmployerPercentAmount = 0
                             rstItems!EmployerTotalAmount = 0
                            rstItems.Update
                            GoTo NEXTGROSS
                        End If
                        
                        'Test Max value
                        ItemTotalAmount = Round(CDbl((rstItems!ItemAmount + (GROSS * (rstItems!ItemPercent / 100)))))
                        
                        If rsMinMax.RecordCount > 0 Then                                                'TestMin/Max
                            rsMinMax.Find "PyrlItemID = '" & rstItems!PyrlItemID & "'"
                            If Not rsMinMax.EOF Then
                                    MaxValue = TestMinMax(rstItems!YTDMax, rsMinMax!TotalAmount, ItemTotalAmount)
                                    Select Case MaxValue
                                        Case Is > 0
                                            ItemTotalAmount = Round(CDbl(MaxValue))   'Item passed YTDMax test as partial amount-Limit was reached
                                        Case False
                                            'rstItems.Edit
                                              rstItems!TotalAmount = 0
                                              rstItems!PercentAmount = 0
                                              rstItems!EmployerPercentAmount = 0
                                              rstItems!EmployerTotalAmount = 0
                                            rstItems.Update
                                            GoTo NEXTGROSS  'Item is over the max
                                    End Select
                            End If
                        End If
                        
                        'rstItems.Edit
                          rstItems!TotalAmount = ItemTotalAmount
                          rstItems!PercentAmount = (GROSS * (rstItems!ItemPercent / 100))
                          rstItems!EmployerPercentAmount = (GROSS * (rstItems!EmployerItemPercent / 100))
                          rstItems!EmployerTotalAmount = Round(CDbl((rstItems!EmployerItemAmount + (GROSS * (rstItems!EmployerItemPercent / 100)))))
                        rstItems.Update
                    
                    Select Case rstItems![Type]
                     Case "Addition"
                            GROSS = GROSS + ItemTotalAmount
                            AGI = AGI + ItemTotalAmount
                            ADDITIONS = ADDITIONS + ItemTotalAmount
                            'deductions from gross
                     Case "State Tax"
                            STATETAX = STATETAX + ItemTotalAmount
                            AGI = AGI - ItemTotalAmount
                            DEDUCTIONS = DEDUCTIONS + ItemTotalAmount
                     Case "Local Tax"
                            LocalTax = LocalTax + ItemTotalAmount
                            AGI = AGI - ItemTotalAmount
                            DEDUCTIONS = DEDUCTIONS + ItemTotalAmount
                     Case "Deduction"
                            GROSSDED = GROSSDED + ItemTotalAmount
                            AGI = AGI - ItemTotalAmount
                            DEDUCTIONS = DEDUCTIONS + ItemTotalAmount
                     End Select
                End If
NEXTGROSS:    rstItems.MoveNext
            Loop
        
        If GROSS <= 0 Then
            GROSSDED = 0
            GROSS = 0
            AGI = 0
        Else
            If GROSSDED > GROSS Then
                GROSSDED = 0
                DEDUCTIONS = DEDUCTIONS - GROSSDED
                AGI = GROSS
            End If
        End If
   End If  'REC COUNT > 0
   
   If AGI > 0 Then      'Calc Taxes
      
       If Me.Caption = "Payroll - Check Detail" Then '1--- Pyrl - Check Detail
            If ckCheckDetail.Value = 0 Then 'Manual check
                FIT = Rst!LASTFIT
                GoTo FICA
            Else
                GoTo CALCFIT
            End If
        End If
            
CALCFIT:    'Calc fit withholding
            If Rst!FITYN.Value = -1 Then
            
               'Open Tax Withholding Table recordset
                Dim rstFIT As ADODB.Recordset
                
                Set rstFIT = New ADODB.Recordset
                rstFIT.Open "SELECT [OneAllowance] FROM [Pyrl - Withholding] WHERE [Period] = '" & Rst![PAYFREQUENCY] & "' AND [OneAllowance] > 0", db, adOpenKeyset, adLockOptimistic, adCmdText
                If rstFIT.RecordCount > 0 Then
                    rstFIT.MoveFirst
                End If
                AllowAmt = (rstFIT!OneAllowance * Rst!FEDALLOW)
                'rstFIT.Find ("[Period] = '" & rst![PAYFREQUENCY] & "' AND [OneAllowance] > 0")
                rstFIT.Close
                Set rstFIT = Nothing
                
                Set rstFIT = New ADODB.Recordset
                'Debug.Print "SELECT [Over],[Bracket],[Cumulative] FROM [Pyrl - Withholding] WHERE [Status] = '" & rst![FEDFILINGSTATUS] & "' AND [Period] = '" & rst![PAYFREQUENCY] & "' AND [Over] < " & (AGI - AllowAmt) & " AND [NotOver] > " & (AGI - AllowAmt)
                rstFIT.Open "SELECT [Over],[Bracket],[Cumulative] FROM [Pyrl - Withholding] WHERE [Status] = '" & Rst![FEDFILINGSTATUS] & "' AND [Period] = '" & Rst![PAYFREQUENCY] & "' AND [Over] < " & (AGI - AllowAmt) & " AND [NotOver] > " & (AGI - AllowAmt), db, adOpenKeyset, adLockOptimistic, adCmdText
                'rstFIT.Find ("[Status] = '" & rst![FEDFILINGSTATUS] & "' AND [Period] = '" & rst![PAYFREQUENCY] & "' AND [Over] < " & (AGI - AllowAmt) & "AND [NotOver] > " & (AGI - AllowAmt))
                If rstFIT.RecordCount > 0 Then
                    rstFIT.MoveFirst
                End If
                If Not rstFIT.EOF Then
                    FIT = ((AGI - AllowAmt - rstFIT!Over) * rstFIT![Bracket]) + rstFIT![Cumulative] + Rst!FEDWITHHOLDAMT
                Else
                    FIT = NZ(Rst!FEDWITHHOLDAMT, 0)
                End If
                
                rstFIT.Close
                Set rstFIT = Nothing
            End If

                    
           
FICA:
        'calc fica
        Medi = 0
        SocSec = 0
    If Me.Caption = "Payroll - Check Detail" Then  '1--- Pyrl - Check Detail
        If ckCheckDetail.Value = 0 Then 'Manual check
            FicaEE = Rst!LASTFICA 'Manual Entry
            FICAER = (FicaEE / (1 - (NZ(rstSETUP![FICA EMPL PERCENT], 0) / 100))) * (NZ(rstSETUP![FICA EMPL PERCENT], 0) / 100)
            GoTo FUTA
        Else
            GoTo CALCFICA
        End If
    End If
        
        
        
CALCFICA:
        If Rst!FICAYN.Value = -1 Then
            If Rst!YTDGROSS < NZ(rstSETUP!SSWAGEBASE, 0) Then
                If (Rst!YTDGROSS + GROSS) <= NZ(rstSETUP!SSWAGEBASE, 0) Then
                    SocSec = GROSS * (NZ(rstSETUP!FICASS, 0) / 100)
                Else
                    SocSec = (NZ(rstSETUP!SSWAGEBASE, 0) - Rst!YTDGROSS) * (NZ(rstSETUP!FICASS, 0) / 100)
                End If
             End If
            
            If NZ(rstSETUP!MEDIWAGEBASE, 0) > 0 Then
                 If Rst!YTDGROSS < NZ(rstSETUP!MEDIWAGEBASE, 0) Then
                    If (Rst!YTDGROSS + GROSS) <= NZ(rstSETUP!MEDIWAGEBASE, 0) Then
                        Medi = GROSS * (rstSETUP!FICAMED / 100)
                     Else
                        Medi = (NZ(rstSETUP!MEDIWAGEBASE, 0) - Rst!YTDGROSS) * (NZ(rstSETUP!FICAMED, 0) / 100)
                    End If
                End If
            Else
                Medi = GROSS * (NZ(rstSETUP!FICAMED, 0) / 100)
            End If
    
            FicaEE = (Medi + SocSec) * (1 - (NZ(rstSETUP![FICA EMPL PERCENT], 0) / 100))
            FICAER = (Medi + SocSec) * (NZ(rstSETUP![FICA EMPL PERCENT], 0) / 100)
        End If
    
        
FUTA:    'Calc FUTA Employer tax
        If Rst!YTDGROSS < NZ(rstSETUP!FUTAWAGEBASE, 0) Then
            If (Rst!YTDGROSS + GROSS) <= NZ(rstSETUP!FUTAWAGEBASE, 0) Then
                FUTA = GROSS * (NZ(rstSETUP!FUTARATE, 0) / 100)
            Else
                FUTA = (NZ(rstSETUP!FUTAWAGEBASE, 0) - Rst!YTDGROSS) * (NZ(rstSETUP!FUTARATE, 0) / 100)
            End If
        End If
        
            'Calc SUI Employer tax
        If Rst!YTDGROSS < NZ(rstSETUP!SUIWAGEBASE, 0) Then
            If (Rst!YTDGROSS + GROSS) <= NZ(rstSETUP!SUIWAGEBASE, 0) Then
                SUI = GROSS * (NZ(rstSETUP!SUIRATE, 0) / 100)
            Else
                SUI = (NZ(rstSETUP!SUIWAGEBASE, 0) - Rst!YTDGROSS) * (NZ(rstSETUP!SUIRATE, 0) / 100)
            End If
        End If
  
End If      ' AGI was zero
    
    FIT = Round(CDbl(FIT))
    FicaEE = Round(CDbl(FicaEE))
        
  If rstItems.RecordCount > 0 Then
        rstItems.MoveFirst              'calc payroll items with basis of adjusted gross
        Do Until rstItems.EOF
            If rstItems![Basis] = "AGI" Then
                    
                    If TestHighLow(Rst!YTDGROSS, rstItems!WageLow, rstItems!WageHigh) = False Then 'Test High/Low
                            'rstItems.Edit
                              rstItems!TotalAmount = 0
                              rstItems!PercentAmount = 0
                              rstItems!EmployerPercentAmount = 0
                              rstItems!EmployerTotalAmount = 0
                            rstItems.Update
                            GoTo NEXTAGI
                        End If
                        
                        'Test Max value
                        ItemTotalAmount = Round(CDbl((rstItems!ItemAmount + (AGI * (rstItems!ItemPercent / 100)))))
                        If rsMinMax.RecordCount > 0 Then                                                'TestMin/Max
                            rsMinMax.Find ("PyrlItemID = '" & rstItems!PyrlItemID & "'")
                            If Not rsMinMax.EOF Then
                                    MaxValue = TestMinMax(rstItems!YTDMax, rsMinMax!TotalAmount, ItemTotalAmount)
                                    Select Case MaxValue
                                        Case Is > 0
                                            ItemTotalAmount = Round(CDbl(MaxValue))   'Item passed YTDMax test as partial amount-Limit was reached
                                        Case False
                                            'rstItems.Edit
                                              rstItems!TotalAmount = 0
                                              rstItems!PercentAmount = 0
                                              rstItems!EmployerPercentAmount = 0
                                              rstItems!EmployerTotalAmount = 0
                                            rstItems.Update
                                            GoTo NEXTAGI  'Item is over the max
                                    End Select
                            End If
                        End If
                       'rstItems.Edit
                         rstItems!TotalAmount = ItemTotalAmount
                         rstItems!PercentAmount = (AGI * (rstItems!ItemPercent / 100))
                         rstItems!EmployerPercentAmount = (AGI * (rstItems!EmployerItemPercent / 100))
                         rstItems!EmployerTotalAmount = Round(CDbl((rstItems!EmployerItemAmount + (AGI * (rstItems!EmployerItemPercent / 100)))))
                       rstItems.Update
                        
                      Select Case rstItems![Type]
                       Case "Addition"
                            GROSS = GROSS + ItemTotalAmount
                            Net = Net + ItemTotalAmount
                            AGIDED = AGIDED + ItemTotalAmount
                            ADDITIONS = ADDITIONS + ItemTotalAmount
                        
                        Case "State Tax"           'deductions from agi
                            If AGI - FicaEE - FIT - ItemTotalAmount >= 0 Then 'Item cannot drive AGI negative
                                STATETAX = STATETAX + ItemTotalAmount
                                STATETAX = IIf(STATETAX < 0, 0, STATETAX)
                                DEDUCTIONS = DEDUCTIONS + ItemTotalAmount
                             Else
                                'rstItems.Edit
                                rstItems![ApplyItem] = 0    'Deselect item that was to large to deduct
                                rstItems.Update
                                GoTo NEXTAGI
                             End If
                        
                        Case "Local Tax"
                            If AGI - FicaEE - FIT - ItemTotalAmount >= 0 Then
                                LocalTax = LocalTax + ItemTotalAmount
                                LocalTax = IIf(LocalTax < 0, 0, LocalTax)
                                DEDUCTIONS = DEDUCTIONS + ItemTotalAmount
                             Else
                                'rstItems.Edit
                                rstItems![ApplyItem] = 0
                                rstItems.Update
                                GoTo NEXTAGI
                             End If
                            
                        Case "Deduction"
                             If AGI - FicaEE - FIT - ItemTotalAmount >= 0 Then
                                DEDUCTIONS = DEDUCTIONS + ItemTotalAmount
                             Else
                                'rstItems.Edit
                                rstItems![ApplyItem] = 0
                                rstItems.Update
                                GoTo NEXTAGI
                             End If
                        
                        End Select
            End If
NEXTAGI:   rstItems.MoveNext
        Loop
    
    End If  'REC COUNT > 0
        
        If AGI < 0 Then
        AGI = 0
        End If

DEDUCTIONS = DEDUCTIONS + FIT + FicaEE  ''
GROSS = Round(CDbl(GROSS))
DEDUCTIONS = Round(CDbl(DEDUCTIONS))
Net = GROSS - DEDUCTIONS  'PreCalc net

' calc net
       
    If rstItems.RecordCount > 0 Then        'calc payroll items with basis of net
        rstItems.MoveFirst
        Do Until rstItems.EOF
            If rstItems![Basis] = "Net" Then
                If TestHighLow(Rst!YTDGROSS, rstItems!WageLow, rstItems!WageHigh) = False Then 'Test High/Low
                            'rstItems.Edit
                              rstItems!TotalAmount = 0
                              rstItems!PercentAmount = 0
                              rstItems!EmployerPercentAmount = 0
                              rstItems!EmployerTotalAmount = 0
                            rstItems.Update
                            GoTo NEXTNET
                        End If
                        
                        'Test Max value
                        ItemTotalAmount = Round(CDbl((rstItems!ItemAmount + (Net * (rstItems!ItemPercent / 100)))))
                        If rsMinMax.RecordCount > 0 Then                                                'TestMin/Max
                            rsMinMax.Find ("PyrlItemID = '" & rstItems!PyrlItemID & "'")
                            If Not rsMinMax.EOF Then
                                    MaxValue = TestMinMax(rstItems!YTDMax, rsMinMax!TotalAmount, ItemTotalAmount)
                                    Select Case MaxValue
                                        Case Is > 0
                                            ItemTotalAmount = Round(CDbl(MaxValue))   'Item passed YTDMax test as partial amount-Limit was reached
                                        Case False
                                            'rstItems.Edit
                                              rstItems!TotalAmount = 0
                                              rstItems!PercentAmount = 0
                                              rstItems!EmployerPercentAmount = 0
                                              rstItems!EmployerTotalAmount = 0
                                            rstItems.Update
                                            GoTo NEXTNET  'Item is over the max
                                    End Select
                            End If
                        End If
                        'rstItems.Edit
                         rstItems!TotalAmount = ItemTotalAmount
                         rstItems!PercentAmount = (Net * (rstItems!ItemPercent / 100))
                         rstItems!EmployerPercentAmount = (Net * (rstItems!EmployerItemPercent / 100))
                         rstItems!EmployerTotalAmount = Round(CDbl((rstItems!EmployerItemAmount + (Net * (rstItems!EmployerItemPercent / 100)))))
                        rstItems.Update
                
                Select Case rstItems![Type]
                Case "Addition"
                    ADDITIONS = ADDITIONS + ItemTotalAmount
                    NETADDITIONS = NETADDITIONS + ItemTotalAmount
                    Net = Net + ItemTotalAmount
                Case Else
                    If Net - ItemTotalAmount >= 0 Then  'Item cannot drive net negative
                        DEDUCTIONS = DEDUCTIONS + ItemTotalAmount
                        Net = Net - ItemTotalAmount
                    Else
                        'rstItems.Edit
                        rstItems![ApplyItem] = 0    'Deselect item that was to large to deduct
                        rstItems.Update
                        GoTo NEXTNET
                    End If
                
                End Select
               
        
            End If      'Basis not = net
NEXTNET:    rstItems.MoveNext
        Loop

End If   'Rec count = 0
        
        GROSS = Round(CDbl(GROSS))
        DEDUCTIONS = Round(CDbl(DEDUCTIONS))
        NETADDITIONS = Round(CDbl(NETADDITIONS))
        Net = GROSS - DEDUCTIONS + NETADDITIONS
        
        If Net < 0 Then
            Net = 0
        End If
        
        PRETAXDED = GROSSDED + AGIDED
    
    GROSS = Round(CDbl(GROSS))
    PRETAXDED = Round(CDbl(PRETAXDED))
    AGI = Round(CDbl(AGI))
    FIT = Round(CDbl(FIT))
    FicaEE = Round(CDbl(FicaEE))
    FICAER = Round(CDbl(FICAER))
    FUTA = Round(CDbl(FUTA))
    SUI = Round(CDbl(SUI))
    STATETAX = Round(CDbl(STATETAX))
    LocalTax = Round(CDbl(LocalTax))
    Net = Round(CDbl(Net))
    ADDITIONS = Round(CDbl(ADDITIONS))
    DEDUCTIONS = Round(CDbl(DEDUCTIONS))

Exit Sub

FormulasError:
ShowStatus False
Call ErrorLog("Pay Employees", "Module-Formulas", Now, Err.Number, Err.Description, True, db)
Resume Next

End Sub

Private Function TestHighLow(YTDGROSS, WageLow, WageHigh)
  
  'Test wage High/Low Parameters
                        If Rst!YTDGROSS >= WageLow Then
                                If Not WageHigh = 0 Then
                                        If YTDGROSS >= WageHigh Then
                                            TestHighLow = False
                                            Exit Function
                                        End If
                                End If
                        Else
                            TestHighLow = False
                            Exit Function
                        End If
               
        TestHighLow = True

End Function

Private Function TestMinMax(YTDMax, SumOfItem, ItemAmt)

'Test Max parameters
                    If YTDMax > 0 Then
                        If YTDMax > SumOfItem Then
                                
                                If YTDMax > SumOfItem + ItemAmt Then
                                    TestMinMax = True
                                    Exit Function
                                Else
                                    TestMinMax = (YTDMax - SumOfItem)
                                    Exit Function
                                End If
                        
                        Else
                            TestMinMax = False
                            Exit Function
                        End If
                    End If
        
        TestMinMax = True
End Function

Private Sub ResetPyrlItems()

'Clear Work Table
db.Execute "Delete * from [Pyrl - Select Pyrl Items Work]"
db.Execute "INSERT INTO [Pyrl - Select Pyrl Items Work] Select * from [Pyrl - Payroll Items]"
End Sub

Private Sub BlankCheck()

'Open Employee Data
    'Set db = CurrentDb
    Dim rstBlankCheck As ADODB.Recordset
   
    
    Set rstBlankCheck = New ADODB.Recordset
    rstBlankCheck.Open "SELECT * FROM [Pyrl - Employees]WHERE [EMP ID] = '" & txtCheckDetail(0).Text & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
    rstBlankCheck.MoveFirst
    
    With rstBlankCheck
    LASTHOURS = !LASTHOURS
    '.Edit
    !PAY = -1
    !LASTGROSS = 0
    !LASTAGI = 0
    !LASTFIT = 0
    !LASTSTATETAX = 0
    !LASTLOCAL = 0
    !LASTFICA = 0
    !NETCHECK = 0
    !LASTHOURS = 0
    !PRETAXDED = 0
    !DEDUCTIONS = 0
    !ADDITIONS = 0
    !Amount = 0
    !BlankCheckHourlyRate = 0
    !BlankCheckOTRate = 0
    .Update
    .Close
    End With
    Set rstBlankCheck = Nothing
End Sub


Private Sub txtPyrllItems_LostFocus(Index As Integer)
Select Case Index
Case 0
    If txtPyrllItems(0).Text <> tempPayItems Then
        LoadPayrollDB
        tempPayItems = txtPyrllItems(0).Text
    End If
End Select
End Sub


