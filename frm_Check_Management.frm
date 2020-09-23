VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_Check_Management 
   Caption         =   "Check Management"
   ClientHeight    =   7710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13425
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7710
   ScaleWidth      =   13425
   Begin VB.Frame frPrimary 
      Height          =   7215
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   12975
      Begin VB.CommandButton cmdBackChk 
         Caption         =   "&Back"
         Height          =   855
         Left            =   11880
         Picture         =   "frm_Check_Management.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   360
         Width           =   975
      End
      Begin VB.Frame Frame3 
         Height          =   1335
         Left            =   7800
         TabIndex        =   17
         Top             =   960
         Width           =   2895
         Begin VB.CommandButton cmdSearch 
            Caption         =   "&Search"
            Height          =   855
            Left            =   1680
            Picture         =   "frm_Check_Management.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtfields 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "MM/dd/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   240
            TabIndex        =   18
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Check No"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   19
            Top             =   480
            Width           =   1335
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1335
         Left            =   4560
         TabIndex        =   12
         Top             =   960
         Width           =   3135
         Begin VB.TextBox txtChkManage 
            Height          =   285
            Index           =   1
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   840
            Width           =   1575
         End
         Begin VB.TextBox txtChkManage 
            Height          =   285
            Index           =   0
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   13
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label lblLabels 
            Caption         =   "Total Payment:"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   16
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblLabels 
            Caption         =   "Checks Total:"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   15
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2055
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   4335
         Begin VB.CommandButton cmdLookupVend 
            Height          =   285
            Left            =   2760
            Picture         =   "frm_Check_Management.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   1560
            Width           =   375
         End
         Begin VB.TextBox txtfields 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "MM/dd/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            Height          =   285
            Index           =   35
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   21
            Top             =   1560
            Width           =   1455
         End
         Begin VB.CommandButton cmdShow 
            Caption         =   "&Execute"
            Height          =   855
            Left            =   3240
            Picture         =   "frm_Check_Management.frx":091E
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   1080
            Width           =   975
         End
         Begin VB.CommandButton cmdDate 
            Height          =   285
            Index           =   1
            Left            =   2760
            Picture         =   "frm_Check_Management.frx":0C28
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   960
            Width           =   375
         End
         Begin VB.TextBox txtfields 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "MM/dd/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   960
            Width           =   1455
         End
         Begin VB.CommandButton cmdDate 
            Height          =   285
            Index           =   0
            Left            =   2760
            Picture         =   "frm_Check_Management.frx":0F32
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   360
            Width           =   375
         End
         Begin VB.TextBox txtfields 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "MM/dd/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label lblLabels 
            Caption         =   "Bank Account:"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   22
            Top             =   1560
            Width           =   1095
         End
         Begin VB.Label lblLabels 
            Caption         =   "End Date:"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   10
            Top             =   960
            Width           =   975
         End
         Begin VB.Label lblLabels 
            Caption         =   "Start Date:"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   9
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   855
         Left            =   11880
         Picture         =   "frm_Check_Management.frx":123C
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1320
         Width           =   975
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   855
         Left            =   10800
         Picture         =   "frm_Check_Management.frx":1546
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1320
         Width           =   975
      End
      Begin MSDataGridLib.DataGrid grdDatagrid 
         Height          =   4695
         Left            =   60
         TabIndex        =   30
         Top             =   2400
         Width           =   12735
         _ExtentX        =   22463
         _ExtentY        =   8281
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
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
         Caption         =   "Posted/Used Check"
         ColumnCount     =   10
         BeginProperty Column00 
            DataField       =   "AP PAY Check No"
            Caption         =   "Check No"
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
            DataField       =   "AP PAY Type"
            Caption         =   "Check Type"
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
            DataField       =   "AP PAY Vendor No"
            Caption         =   "Vendor ID"
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
            DataField       =   "AP PAY Transaction Date"
            Caption         =   "Transaction Date"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "d MMMM, yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "AP PAY Amount"
            Caption         =   "Check Amount"
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
            DataField       =   "AP PAY Bank Account"
            Caption         =   "Bank Account"
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
            DataField       =   "AP PAY Status"
            Caption         =   "Check Status"
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
            DataField       =   "AP PAY Void"
            Caption         =   "Void"
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
         BeginProperty Column08 
            DataField       =   "AP PAY Posted YN"
            Caption         =   "Posted"
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
         BeginProperty Column09 
            DataField       =   "AP PAY Cleared"
            Caption         =   "Cleared"
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
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            BeginProperty Column00 
               DividerStyle    =   1
               ColumnWidth     =   1244.976
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1544.882
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1454.74
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1319.811
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               ColumnWidth     =   1440
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1544.882
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   794.835
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   854.929
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   824.882
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame frDua 
      Height          =   7215
      Left            =   0
      TabIndex        =   31
      Top             =   480
      Visible         =   0   'False
      Width           =   11655
      Begin VB.Frame Frame6 
         Height          =   2055
         Left            =   120
         TabIndex        =   44
         Top             =   240
         Width           =   4335
         Begin VB.TextBox txtfields 
            DataField       =   "AR SALE Ship Date"
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
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   51
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton cmdDate 
            Height          =   285
            Index           =   3
            Left            =   2760
            Picture         =   "frm_Check_Management.frx":1850
            Style           =   1  'Graphical
            TabIndex        =   50
            Top             =   360
            Width           =   375
         End
         Begin VB.TextBox txtfields 
            DataField       =   "AR SALE Ship Date"
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
            Index           =   5
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   49
            Top             =   960
            Width           =   1455
         End
         Begin VB.CommandButton cmdDate 
            Height          =   285
            Index           =   2
            Left            =   2760
            Picture         =   "frm_Check_Management.frx":1E2A
            Style           =   1  'Graphical
            TabIndex        =   48
            Top             =   960
            Width           =   375
         End
         Begin VB.CommandButton cmdExecuteBooked 
            Caption         =   "&Execute"
            Height          =   855
            Left            =   3240
            Picture         =   "frm_Check_Management.frx":2404
            Style           =   1  'Graphical
            TabIndex        =   47
            Top             =   1080
            Width           =   975
         End
         Begin VB.TextBox txtfields 
            DataField       =   "AR SALE Ship Date"
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
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   46
            Top             =   1560
            Width           =   1455
         End
         Begin VB.CommandButton Command1 
            Height          =   285
            Left            =   2760
            Picture         =   "frm_Check_Management.frx":270E
            Style           =   1  'Graphical
            TabIndex        =   45
            Top             =   1560
            Width           =   375
         End
         Begin VB.Label lblLabels 
            Caption         =   "Start Date:"
            Height          =   255
            Index           =   12
            Left            =   120
            TabIndex        =   54
            Top             =   360
            Width           =   855
         End
         Begin VB.Label lblLabels 
            Caption         =   "End Date:"
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   53
            Top             =   960
            Width           =   975
         End
         Begin VB.Label lblLabels 
            Caption         =   "Bank Account:"
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   52
            Top             =   1560
            Width           =   1095
         End
      End
      Begin VB.Frame Frame5 
         Height          =   2055
         Left            =   4560
         TabIndex        =   39
         Top             =   240
         Width           =   1815
         Begin VB.TextBox txtChkManage 
            Height          =   285
            Index           =   3
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   41
            Top             =   720
            Width           =   1575
         End
         Begin VB.TextBox txtChkManage 
            Height          =   285
            Index           =   2
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   40
            Top             =   1440
            Width           =   1575
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Checks Total"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   43
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Total Payment"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   42
            Top             =   1200
            Width           =   1575
         End
      End
      Begin VB.CommandButton cmdPrintBooked 
         Caption         =   "&Print"
         Height          =   855
         Left            =   9480
         Picture         =   "frm_Check_Management.frx":2858
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   1320
         Width           =   975
      End
      Begin VB.CommandButton cmdRefreshBooked 
         Caption         =   "&Refresh"
         Height          =   855
         Left            =   10560
         Picture         =   "frm_Check_Management.frx":2B62
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   1320
         Width           =   975
      End
      Begin VB.Frame Frame4 
         Height          =   1335
         Left            =   6480
         TabIndex        =   33
         Top             =   960
         Width           =   2895
         Begin VB.TextBox txtfields 
            DataField       =   "AR SALE Ship Date"
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
            Index           =   3
            Left            =   240
            TabIndex        =   35
            Top             =   720
            Width           =   1335
         End
         Begin VB.CommandButton cmdSearchBooked 
            Caption         =   "&Search"
            Height          =   855
            Left            =   1680
            Picture         =   "frm_Check_Management.frx":2E6C
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   360
            Width           =   975
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Check No"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   7
            Left            =   240
            TabIndex        =   36
            Top             =   480
            Width           =   1335
         End
      End
      Begin VB.CommandButton cmdBackBooked 
         Caption         =   "&Back"
         Height          =   855
         Left            =   10560
         Picture         =   "frm_Check_Management.frx":3176
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   360
         Width           =   975
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frm_Check_Management.frx":3480
         Height          =   4695
         Left            =   120
         TabIndex        =   55
         Top             =   2400
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   8281
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
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
         Caption         =   "Booked Check "
         ColumnCount     =   9
         BeginProperty Column00 
            DataField       =   "AP PO Check Number"
            Caption         =   "Check No."
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
            DataField       =   "AP PO Amount Paid"
            Caption         =   "Amount Paid"
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
         BeginProperty Column02 
            DataField       =   "AP PO Ext Document No"
            Caption         =   "Doc. No"
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
            DataField       =   "AP PO Total Amount"
            Caption         =   "Total Amount"
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
         BeginProperty Column04 
            DataField       =   "AP PO Vendor Name"
            Caption         =   "Vendor Name"
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
            DataField       =   "AP PO Due Date"
            Caption         =   "Due Date"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "MM/dd/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "AP PO Vendor Invoice No"
            Caption         =   "Invoice No"
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
            DataField       =   "AP PO Check Date"
            Caption         =   "Check Date"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "MM/dd/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "AP PO Check Acct ID"
            Caption         =   "Check Account"
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
            MarqueeStyle    =   3
            BeginProperty Column00 
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column01 
               Alignment       =   1
               ColumnWidth     =   1124.787
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1170.142
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               ColumnWidth     =   1230.236
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1544.882
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1124.787
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1335.118
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame frUtama 
      Height          =   3375
      Left            =   5040
      TabIndex        =   24
      Top             =   480
      Width           =   5055
      Begin VB.CommandButton cmdPosted 
         Caption         =   "Posted/Check"
         Height          =   495
         Left            =   840
         TabIndex        =   26
         Top             =   1800
         Width           =   1335
      End
      Begin VB.CommandButton cmdBook 
         Caption         =   "Booked"
         Height          =   495
         Left            =   840
         TabIndex        =   25
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Description"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2520
         TabIndex        =   28
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Height          =   2655
         Left            =   2520
         TabIndex        =   27
         Top             =   600
         Width           =   2415
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Check Management"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   12945
   End
End
Attribute VB_Name = "frm_Check_Management"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents ADOprimaryrs As ADODB.Recordset
Attribute ADOprimaryrs.VB_VarHelpID = -1
Dim db As ADODB.Connection
Dim TempStr As String
Dim WhichField As String

Private Sub cmdBackChk_Click()
    frPrimary.Visible = False
    frUtama.ZOrder 0
    frUtama.Visible = True
    Form_Resize
End Sub

Private Sub cmdBook_Click()
    txtfields(0) = ""
    txtfields(1) = ""
    txtfields(5) = ""
    txtfields(6) = ""
    WhichField = ""
    frPrimary.Visible = False
    frUtama.Visible = False
    frDua.ZOrder 0
    frDua.Visible = True
    OpenDB
    Form_Resize
End Sub

Private Sub cmdBook_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label2.Caption = "Booked but not Posted"
End Sub

Private Sub cmdDate_Click(Index As Integer)
Select Case Index
Case 0, 3
    Menu_Calendar.WhoCallMe True, 1302
    'Menu_Calendar.Show vbModal
    txtfields(6) = txtfields(0)
Case 1, 2
    Menu_Calendar.WhoCallMe True, 1640
    'Menu_Calendar.Show vbModal
    txtfields(5) = txtfields(1)
End Select
End Sub

Private Sub cmdExecuteBooked_Click()
    If txtfields(6).Text = "" Or txtfields(5) = "" Then
        MsgBox "You must complete the start and the end date before you could continue", vbInformation, "Information"
        Exit Sub
    End If
    OpenDB
End Sub

Private Sub cmdLookupVend_Click()
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
    txtfields(4) = txtfields(35)
End Sub

Private Sub cmdPosted_Click()
    txtfields(0) = ""
    txtfields(1) = ""
    txtfields(5) = ""
    txtfields(6) = ""
    WhichField = ""
    frPrimary.Visible = True
    frPrimary.ZOrder 0
    frUtama.Visible = False
    OpenDB
    Form_Resize
End Sub

Private Sub cmdPosted_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label2.Caption = "Posted Check"
End Sub

Private Sub cmdRefresh_Click()

ShowStatus True
    txtfields(0) = ""
    txtfields(1) = ""
    txtfields(35) = ""
    OpenDB
ShowStatus False
End Sub

Private Sub cmdRefreshBooked_Click()
ShowStatus True
    txtfields(4) = ""
    txtfields(5) = ""
    txtfields(6) = ""
    OpenDB
ShowStatus False
End Sub

Private Sub cmdSearchBooked_Click()
On Error GoTo NOTFOUND
    If txtfields(3) = "" Then Exit Sub
    DataGrid1.SetFocus
    ADOprimaryrs.MoveFirst
    
    If WhichField = "" Then
        WhichField = "AP PO Check Number"
    End If
    If ADOprimaryrs("" & WhichField & "").Type = 202 Then
        ADOprimaryrs.Find "[" & WhichField & "]='" & txtfields(3) & "'"
    Else
        ADOprimaryrs.Find "[" & WhichField & "]=" & txtfields(3)
    End If
    
    If ADOprimaryrs.EOF Then
NOTFOUND:
        MsgBox lblLabels(7) & " " & txtfields(3) & " is not existed.", vbInformation, "Information"
    End If
    SendKeys ("{LEFT}")
End Sub

Private Sub cmdSearch_Click()
On Error GoTo NOTFOUND
    If txtfields(2) = "" Then Exit Sub
    grdDataGrid.SetFocus
    ADOprimaryrs.MoveFirst
    
    If WhichField = "" Then
        WhichField = "AP PAY Check No"
    End If
    If ADOprimaryrs("" & WhichField & "").Type = 202 Then
        ADOprimaryrs.Find "[" & WhichField & "]='" & txtfields(2) & "'"
    Else
        ADOprimaryrs.Find "[" & WhichField & "]=" & txtfields(2)
    End If
    
    If ADOprimaryrs.EOF Then
NOTFOUND:
        MsgBox lblLabels(5) & " " & txtfields(2) & " is not existed.", vbInformation, "Information"
    End If
    SendKeys ("{LEFT}")
End Sub

Private Sub cmdShow_Click()
    If txtfields(0).Text = "" Or txtfields(1) = "" Then
        MsgBox "You must complete the start and the end date before you could continue", vbInformation, "Information"
        Exit Sub
    End If
    OpenDB
End Sub

Private Sub cmdBackBooked_Click()
    frDua.Visible = False
    frDua.ZOrder 0
    frUtama.Visible = True
    Form_Resize
End Sub

Private Sub Command1_Click()
    cmdLookupVend_Click
End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
If ADOprimaryrs.RecordCount = 0 Then Exit Sub
    lblLabels(7) = DataGrid1.Columns(ColIndex).Caption
    WhichField = DataGrid1.Columns(ColIndex).DataField
    ADOprimaryrs.Close
    Set ADOprimaryrs = Nothing
    Set ADOprimaryrs = New ADODB.Recordset
    ADOprimaryrs.Open TempStr & " ORDER BY [" & WhichField & "]", db, adOpenKeyset, adLockOptimistic, adCmdText
    Set DataGrid1.DataSource = ADOprimaryrs
End Sub

Private Sub Form_Load()
ShowStatus True
  Set db = New ADODB.Connection
  db.CursorLocation = adUseClient
  db.Open gblADOProvider
    
  GetTextColor Me
ShowStatus False
End Sub

Private Sub OpenDB()
Dim TotalAmount As Currency


If ADOprimaryrs Is Nothing Then
Else
    ADOprimaryrs.Close
    Set ADOprimaryrs = Nothing
End If

  TotalAmount = 0
  
  Set grdDataGrid.DataSource = Nothing
  
  
  Set ADOprimaryrs = New ADODB.Recordset
  
  If frPrimary.Visible = True Then
    TempStr = "SELECT * FROM [AP PAYMENT HEADER] "
    If txtfields(0) = "" And txtfields(1) = "" Then
    Else
      If txtfields(35) = "" Then
        TempStr = TempStr & "WHERE [AP PAY Transaction Date] BETWEEN #" & txtfields(0).Text & "# AND #" & txtfields(1).Text & "#"
      Else
        TempStr = TempStr & "WHERE [AP PAY BANK Account]='" & txtfields(35) & "' AND [AP PAY Transaction Date] BETWEEN #" & txtfields(0).Text & "# AND #" & txtfields(1).Text & "#"
      End If
   End If
        ADOprimaryrs.Open TempStr, db, adOpenKeyset, adLockReadOnly, adCmdText
        
        If ADOprimaryrs.RecordCount > 0 Then
          ADOprimaryrs.MoveFirst
          Do While ADOprimaryrs.EOF = False
            If ADOprimaryrs![AP PAY Void] = False Then
              TotalAmount = TotalAmount + ADOprimaryrs![AP PAY Amount]
            End If
            ADOprimaryrs.MoveNext
          Loop
        End If
        Set grdDataGrid.DataSource = ADOprimaryrs
        txtChkManage(0) = FormatCurr(TotalAmount)
        txtChkManage(1) = ADOprimaryrs.RecordCount
  
  ElseIf frDua.Visible = True Then
    
    TempStr = "SELECT [AP PO Check Number],[AP PO Amount Paid],[AP PO Ext Document No]," & _
    "[AP PO Total Amount],[AP PO Vendor Name],[AP PO Due Date],[AP PO Vendor Invoice No]," & _
    "[AP PO Check Date],[AP PO Check Acct ID],[AP PO Posted YN] FROM [AP Purchase] " & _
    "WHERE [AP PO Posted YN]=False AND [AP PO Check Number] is not null "
    
    If txtfields(5) = "" And txtfields(6) = "" Then
    Else
      If txtfields(4) = "" Then
        TempStr = TempStr & "AND [AP PO Check Date] BETWEEN #" & txtfields(6).Text & "# AND #" & txtfields(5).Text & "#"
      Else
        TempStr = TempStr & "AND [AP PO Check Acct ID]='" & txtfields(4) & "' AND [AP PO Check Date] BETWEEN #" & txtfields(6).Text & "# AND #" & txtfields(5).Text & "#"
      End If
   End If
    
    ADOprimaryrs.Open TempStr, db, adOpenKeyset, adLockReadOnly, adCmdText
        TotalAmount = 0
        If ADOprimaryrs.RecordCount > 0 Then
          ADOprimaryrs.MoveFirst
          Do While ADOprimaryrs.EOF = False
              TotalAmount = TotalAmount + ADOprimaryrs![AP PO Amount Paid]
              ADOprimaryrs.MoveNext
          Loop
        End If
        Set DataGrid1.DataSource = ADOprimaryrs
        txtChkManage(3) = FormatCurr(TotalAmount)
        txtChkManage(2) = ADOprimaryrs.RecordCount
  End If

End Sub

Public Sub OpenPosted(CheckNo As String)
    Me.Show
    cmdPosted_Click
    txtfields(2).Text = CheckNo
    cmdSearch_Click
End Sub

Private Sub Form_Resize()
  If fMainForm.WindowState = 1 Then Exit Sub
  If Me.WindowState = 0 Then
  ElseIf Me.WindowState = 2 Then
    GoTo SkipResize
  Else
    Exit Sub
  End If
  
If frUtama.Visible = True Then
    Me.Width = 5190
    Me.Height = 4260
ElseIf frPrimary.Visible = True Then
    Me.Width = 13155
    Me.Height = 8100
ElseIf frDua.Visible = True Then
    Me.Width = 11790
    Me.Height = 8100
End If
SkipResize:
If frUtama.Visible = True Then
  frUtama.Top = 480
  frUtama.Left = (Me.ScaleWidth - frUtama.Width) / 2
  Label1(2).Width = frUtama.Width
  Label1(2).Left = frUtama.Left
  frUtama.Top = (Me.ScaleHeight - frUtama.Height) / 2 + 230
ElseIf frPrimary.Visible = True Then
  frPrimary.Top = 480
  frPrimary.Left = (Me.ScaleWidth - frPrimary.Width) / 2
  Label1(2).Width = frPrimary.Width
  Label1(2).Left = frPrimary.Left
  frPrimary.Top = (Me.ScaleHeight - frPrimary.Height) / 2 + 230
ElseIf frDua.Visible = True Then
  frDua.Top = 480
  frDua.Left = (Me.ScaleWidth - frDua.Width) / 2
  Label1(2).Width = frDua.Width
  Label1(2).Left = frDua.Left
  frDua.Top = (Me.ScaleHeight - frDua.Height) / 2 + 230
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm_Check_Management = Nothing
End Sub

Private Sub grdDataGrid_HeadClick(ByVal ColIndex As Integer)
If ADOprimaryrs.RecordCount = 0 Then Exit Sub
    lblLabels(5) = grdDataGrid.Columns(ColIndex).Caption
    WhichField = grdDataGrid.Columns(ColIndex).DataField
    ADOprimaryrs.Close
    Set ADOprimaryrs = Nothing
    Set ADOprimaryrs = New ADODB.Recordset
    ADOprimaryrs.Open TempStr & " ORDER BY [" & WhichField & "]", db, adOpenKeyset, adLockOptimistic, adCmdText
    Set grdDataGrid.DataSource = ADOprimaryrs
End Sub

Private Sub txtfields_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case Index
Case 2
    If KeyCode = 13 Then cmdSearch_Click
End Select
End Sub
