VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_Bank_Reconciliation 
   Caption         =   "Bank Reconciliation"
   ClientHeight    =   8670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9900
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8670
   ScaleWidth      =   9900
   Begin VB.Frame frPrimary 
      Height          =   8175
      Left            =   0
      TabIndex        =   12
      Top             =   480
      Width           =   9855
      Begin MSDataGridLib.DataGrid grdDataGrid 
         Height          =   2895
         Left            =   240
         TabIndex        =   19
         Top             =   2160
         Width           =   9465
         _ExtentX        =   16695
         _ExtentY        =   5106
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
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "BANK RECD Doc #"
            Caption         =   "Doc. No."
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
            DataField       =   "BANK RECD Cleared"
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
         BeginProperty Column02 
            DataField       =   "BANK RECD Date"
            Caption         =   "Date"
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
         BeginProperty Column03 
            DataField       =   "BANK RECD Description"
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
         BeginProperty Column04 
            DataField       =   "BANK RECD Type"
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
         BeginProperty Column05 
            DataField       =   "BANK RECD Amount"
            Caption         =   "Amount"
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
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   4
            BeginProperty Column00 
               ColumnWidth     =   1035.213
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   794.835
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1440
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   2849.953
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1335.118
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1349.858
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame3 
         Height          =   1335
         Left            =   4080
         TabIndex        =   23
         Top             =   120
         Width           =   2895
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
            Left            =   240
            TabIndex        =   25
            Top             =   720
            Width           =   1335
         End
         Begin VB.CommandButton cmdSearch 
            Caption         =   "&Search"
            Height          =   855
            Left            =   1680
            Picture         =   "frm_Bank_Reconciliation.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   360
            Width           =   855
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Doc. No."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   24
            Left            =   240
            TabIndex        =   26
            Top             =   480
            Width           =   1335
         End
      End
      Begin VB.CommandButton cmdRecalc 
         Caption         =   "Recalc"
         Height          =   375
         Left            =   8520
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdPost 
         Caption         =   "Post"
         Height          =   375
         Left            =   7080
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
      Begin VB.Frame Frame1 
         Height          =   855
         Left            =   7080
         TabIndex        =   16
         Top             =   600
         Width           =   2655
         Begin VB.OptionButton optType 
            Caption         =   "Credit"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   18
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton optType 
            Caption         =   "Debit"
            Height          =   255
            Index           =   1
            Left            =   1320
            TabIndex        =   17
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.TextBox txtLabels 
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
         Index           =   11
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1080
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Frame Frame4 
         Height          =   3015
         Left            =   120
         TabIndex        =   27
         Top             =   5040
         Width           =   9615
         Begin VB.CommandButton cmdDate 
            Height          =   285
            Index           =   1
            Left            =   2640
            Picture         =   "frm_Bank_Reconciliation.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   49
            Top             =   1320
            Width           =   375
         End
         Begin VB.CommandButton cmdDate 
            Height          =   285
            Index           =   0
            Left            =   2640
            Picture         =   "frm_Bank_Reconciliation.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   48
            Top             =   840
            Width           =   375
         End
         Begin VB.TextBox txtLabels 
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
            Index           =   10
            Left            =   8040
            Locked          =   -1  'True
            TabIndex        =   47
            Top             =   2520
            Width           =   1455
         End
         Begin VB.TextBox txtLabels 
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
            Index           =   9
            Left            =   8040
            Locked          =   -1  'True
            TabIndex        =   46
            Top             =   2160
            Width           =   1455
         End
         Begin VB.TextBox txtLabels 
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
            Index           =   8
            Left            =   8040
            Locked          =   -1  'True
            TabIndex        =   45
            Top             =   1800
            Width           =   1455
         End
         Begin VB.TextBox txtLabels 
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
            Index           =   7
            Left            =   8040
            Locked          =   -1  'True
            TabIndex        =   44
            Top             =   1440
            Width           =   1455
         End
         Begin VB.TextBox txtLabels 
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
            Index           =   6
            Left            =   8040
            Locked          =   -1  'True
            TabIndex        =   43
            Top             =   1080
            Width           =   1455
         End
         Begin VB.TextBox txtLabels 
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
            Index           =   5
            Left            =   8040
            Locked          =   -1  'True
            TabIndex        =   42
            Top             =   720
            Width           =   1455
         End
         Begin VB.TextBox txtLabels 
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
            Index           =   4
            Left            =   8040
            Locked          =   -1  'True
            TabIndex        =   41
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox txtLabels 
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
            Index           =   3
            Left            =   4800
            Locked          =   -1  'True
            TabIndex        =   40
            Top             =   1800
            Width           =   1455
         End
         Begin VB.TextBox txtLabels 
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
            Index           =   2
            Left            =   4800
            Locked          =   -1  'True
            TabIndex        =   39
            Top             =   1440
            Width           =   1455
         End
         Begin VB.TextBox txtLabels 
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
            Index           =   1
            Left            =   4800
            Locked          =   -1  'True
            TabIndex        =   38
            Top             =   1080
            Width           =   1455
         End
         Begin VB.CommandButton cmdBankAcctCharges 
            Height          =   285
            Left            =   1320
            Picture         =   "frm_Bank_Reconciliation.frx":091E
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   2400
            Width           =   375
         End
         Begin VB.CommandButton cmdbankAcctEarned 
            Height          =   285
            Left            =   4560
            Picture         =   "frm_Bank_Reconciliation.frx":0C28
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   2400
            Width           =   375
         End
         Begin VB.TextBox txtLabels 
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
            Index           =   0
            Left            =   4800
            Locked          =   -1  'True
            TabIndex        =   35
            Top             =   720
            Width           =   1455
         End
         Begin VB.TextBox txtfields 
            DataField       =   "BANK REC Interest Acct"
            Height          =   285
            Index           =   2
            Left            =   3360
            Locked          =   -1  'True
            TabIndex        =   34
            Top             =   2400
            Width           =   1215
         End
         Begin VB.TextBox txtfields 
            DataField       =   "BANK REC Ending Balance"
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
            Index           =   3
            Left            =   1560
            TabIndex        =   33
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox txtfields 
            DataField       =   "BANK REC Cutoff Date"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "M/d/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   285
            Index           =   4
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   32
            Top             =   1320
            Width           =   1095
         End
         Begin VB.TextBox txtfields 
            DataField       =   "BANK REC Start Date"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "M/d/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   285
            Index           =   5
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   31
            Top             =   840
            Width           =   1095
         End
         Begin VB.TextBox txtfields 
            DataField       =   "BANK REC Interest"
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
            Index           =   6
            Left            =   4920
            TabIndex        =   30
            Top             =   2400
            Width           =   1335
         End
         Begin VB.TextBox txtfields 
            DataField       =   "BANK REC Service Acct"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            Height          =   285
            Index           =   12
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   29
            Top             =   2400
            Width           =   1215
         End
         Begin VB.TextBox txtfields 
            DataField       =   "BANK REC Service Charge"
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
            Index           =   8
            Left            =   1680
            TabIndex        =   28
            Top             =   2400
            Width           =   1335
         End
         Begin VB.Label lblLabels 
            Caption         =   "End Book Balance"
            Height          =   255
            Index           =   19
            Left            =   6600
            TabIndex        =   68
            Top             =   2520
            Width           =   1455
         End
         Begin VB.Label lblLabels 
            Caption         =   "Unreconciled"
            Height          =   255
            Index           =   18
            Left            =   6600
            TabIndex        =   67
            Top             =   2160
            Width           =   1455
         End
         Begin VB.Label lblLabels 
            Caption         =   "Stmt Balance"
            Height          =   255
            Index           =   17
            Left            =   6600
            TabIndex        =   66
            Top             =   1800
            Width           =   1455
         End
         Begin VB.Label lblLabels 
            Caption         =   "Bank Balance"
            Height          =   255
            Index           =   16
            Left            =   6600
            TabIndex        =   65
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label lblLabels 
            Caption         =   "-  Credit O/S"
            Height          =   255
            Index           =   15
            Left            =   6600
            TabIndex        =   64
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label lblLabels 
            Caption         =   "Adj. Book Balance"
            Height          =   255
            Index           =   14
            Left            =   6600
            TabIndex        =   63
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label lblLabels 
            Caption         =   "+ Debit O/S"
            Height          =   255
            Index           =   13
            Left            =   6600
            TabIndex        =   62
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label lblLabels 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   12
            Left            =   120
            TabIndex        =   61
            Top             =   2670
            Width           =   2895
         End
         Begin VB.Label lblLabels 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   6
            Left            =   3360
            TabIndex        =   60
            Top             =   2670
            Width           =   2895
         End
         Begin VB.Label lblLabels 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Bank Charges"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   59
            Top             =   2160
            Width           =   2895
         End
         Begin VB.Label lblLabels 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "+ Interest Earned"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   11
            Left            =   3360
            TabIndex        =   58
            Top             =   2160
            Width           =   2895
         End
         Begin VB.Label lblLabels 
            Caption         =   "-  Total Debit:"
            Height          =   255
            Index           =   10
            Left            =   3360
            TabIndex        =   57
            Top             =   1800
            Width           =   1335
         End
         Begin VB.Label lblLabels 
            Caption         =   "+ Total Credit:"
            Height          =   255
            Index           =   9
            Left            =   3360
            TabIndex        =   56
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label lblLabels 
            Caption         =   "Book:"
            Height          =   255
            Index           =   8
            Left            =   3360
            TabIndex        =   55
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label lblLabels 
            Caption         =   "Bank:"
            Height          =   255
            Index           =   7
            Left            =   3360
            TabIndex        =   54
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label lblLabels 
            Caption         =   "Statement Balance:"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   53
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label lblLabels 
            Caption         =   "Stmt End Date:"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   52
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label lblLabels 
            Caption         =   "Stmt Start Date:"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   51
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000001&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Starting Balance"
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   0
            Left            =   3360
            TabIndex        =   50
            Top             =   240
            Width           =   2865
         End
      End
      Begin VB.Frame Frame2 
         Height          =   615
         Left            =   240
         TabIndex        =   69
         Top             =   1440
         Width           =   9495
         Begin VB.CommandButton cmdBankAcct 
            Height          =   285
            Left            =   2640
            Picture         =   "frm_Bank_Reconciliation.frx":0F32
            Style           =   1  'Graphical
            TabIndex        =   75
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txtfields 
            DataField       =   "BANK REC Bank Acct"
            Height          =   285
            Index           =   1
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   74
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox txtLabels 
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
            Index           =   12
            Left            =   4560
            Locked          =   -1  'True
            TabIndex        =   71
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txtLabels 
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
            Index           =   13
            Left            =   7800
            Locked          =   -1  'True
            TabIndex        =   70
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label lblLabels 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Bank Account:"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   76
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label lblLabels 
            Caption         =   "Credits Cleared"
            Height          =   255
            Index           =   21
            Left            =   3360
            TabIndex        =   73
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lblLabels 
            Caption         =   "Debits Cleared"
            Height          =   255
            Index           =   22
            Left            =   6480
            TabIndex        =   72
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Label lblLabels 
         Caption         =   "Period Balace:"
         Height          =   255
         Index           =   20
         Left            =   8400
         TabIndex        =   21
         Top             =   840
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   23
         Left            =   240
         TabIndex        =   20
         Top             =   1110
         Width           =   3135
      End
   End
   Begin VB.PictureBox picStatBox 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   120
      ScaleHeight     =   300
      ScaleWidth      =   12600
      TabIndex        =   6
      Top             =   8520
      Visible         =   0   'False
      Width           =   12600
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Picture         =   "frm_Bank_Reconciliation.frx":123C
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Picture         =   "frm_Bank_Reconciliation.frx":157E
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Picture         =   "frm_Bank_Reconciliation.frx":18C0
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Picture         =   "frm_Bank_Reconciliation.frx":1C02
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   11
         Top             =   0
         Width           =   3360
      End
   End
   Begin VB.PictureBox picButtons 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   12600
      TabIndex        =   0
      Top             =   8520
      Visible         =   0   'False
      Width           =   12600
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   300
         Left            =   4320
         TabIndex        =   2
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   300
         Left            =   3240
         TabIndex        =   3
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   300
         Left            =   2160
         TabIndex        =   4
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   300
         Left            =   1080
         TabIndex        =   1
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   300
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bank Reconciliation"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   7305
   End
End
Attribute VB_Name = "frm_Bank_Reconciliation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ADOprimaryrs As ADODB.Recordset
Attribute ADOprimaryrs.VB_VarHelpID = -1
Dim adoPrimaryRS2 As ADODB.Recordset
Attribute adoPrimaryRS2.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean

Dim db As ADODB.Connection
Dim WhichField As String
Dim TempStr As String

Private Sub FormLoadData()

  db.Execute "DELETE * FROM [Bank Reconciliation]"
  db.Execute "DELETE * FROM [Bank Reconciliation Detail]"
  Dim sAccount As String
  'On Error Resume Next
    'Add a new record and requery
    'Dim rs As ADODB.Recordset
    'Dim rs2 As ADODB.Recordset
    Set ADOprimaryrs = New ADODB.Recordset
    ADOprimaryrs.Open "[BANK Reconciliation]", db, adOpenKeyset, adLockOptimistic, adCmdTable
    If txtFields(1).Text = "" Then
        sAccount = NZ(LookRecord("[GL COA Account No]", "[GL Chart Of Accounts]", db, "[GL COA Asset Type] = 'Cash'"))
    Else
        sAccount = txtFields(1).Text
    End If
    ADOprimaryrs.AddNew
      If Len(sAccount) > 0 Then
        ADOprimaryrs("BANK REC Bank Acct") = sAccount
      End If
    'Me.Requery
    'DoCmd.GoToRecord A_FORM, "Bank Reconciliation", A_FIRST

  'Dim Temp As Variant
  'Forms(Me.Name).Visible = False
  'Temp = glrScaleForm(Me, 640, 480)
  'Call CenterForm(Me)
  'Forms(Me.Name).Visible = True
  If txtFields(4) = "" Then
    ADOprimaryrs![BANK REC Cutoff Date] = Format(Now(), "Short Date")
    ADOprimaryrs![BANK REC Start Date] = Format(DateAdd("d", -30, Now()), "Short Date")
  End If
  'txtFields(4) = Format(Now(), "Short Date")
  'txtFields(5) = Format(DateAdd("d", -30, Now()), "Short Date")
  'Me.Refresh
  
  
  Exit Sub
Form_Open_Error:
  Call ErrorLog("Bank Reconciliation", "Form_Open", Now, Err.Number, Err.Description, True, db)
  Resume Next

End Sub

Private Sub cmdBankAcct_Click()
    Dim SQLstatement As String
    Dim No As Integer
    Dim ghead As String
    Dim fhead As String
    
    No = 1499
    SQLstatement = "select [BANK ACCT ID], [BANK ACCT Name]" & _
                    "from [Bank Accounts]"
    ghead = "Account Description"
    fhead = "Account No//Description"
    
    AllLookup.GetWhichTable No, SQLstatement, ghead, fhead, db
    'MsgBox txtfields(1)
    For Each Ctrl In Me.txtFields
        If Ctrl.Index <> 1 Then Ctrl.Text = ""
    Next
    
    'AllLookup.Show vbModal
    txtFields(1).SetFocus
    'MsgBox txtfields(1)
    'AllLookup.Hide
    ShowStatus True
    LoadEverything
    ShowStatus False
End Sub

Private Sub cmdBankAcctCharges_Click()
   LookupCOA 1490
   txtFields(12).SetFocus
End Sub

Private Sub cmdbankAcctEarned_Click()
   LookupCOA 1480
   txtFields(2).SetFocus
End Sub

Private Sub LookupCOA(WhichButt As Integer)
    Dim SQLstatement As String
    Dim No As Integer
    Dim ghead As String
    Dim fhead As String

    No = WhichButt
    SQLstatement = "select [GL COA Account No], [GL COA Account Name]" & _
                    "from [GL Chart Of Accounts]"
    ghead = "Account Description"
    fhead = "Account No//Description"
    
    AllLookup.GetWhichTable No, SQLstatement, ghead, fhead, db
    'AllLookup.Show vbModal

End Sub

Private Sub cmdDate_Click(Index As Integer)
Select Case Index
Case 0
    Menu_Calendar.WhoCallMe True, 1422
    txtFields(4) = FormatDate(DateAdd("m", 1, txtFields(5)))
    'Menu_Calendar.Show vbModal
    txtFields(4).SetFocus
Case 1
    Menu_Calendar.WhoCallMe True, 1432
    'Menu_Calendar.Show vbModal
    txtFields(4).SetFocus
    'txtFields(4) = txtFields(4)
End Select
    ADOprimaryrs.Update
    LoadnCalculate
End Sub
Private Sub LoadnCalculate()
  
  'FormLoadData
  Call InitForm
  Call LoadDetail
  
  Call CalcRecTotals
  'Call ReconcileDetail_AfterUpdate
  
  'Set grdDataGrid.DataSource = adoPrimaryRS2

End Sub

Private Sub cmdPost_Click()

  'On Error GoTo cmdPost_Error

  'Post this transaction to the general ledger
   Dim Success%

  'Force record save
  'DoCmd.RunMacro "Save Record"
'Debug.Print "SAVE"
  If txtLabels(9) <> "$0.00" Then
    MsgBox "Cannot post with difference not equal to zero!", , "Error"
    Exit Sub
  End If

  ShowStatus True

  db.BeginTrans
  Success% = PostReconciliation(db)
  If Success% = False Then
    db.RollbackTrans
    MsgBox "Transaction NOT Posted."
    'Debug.Print "%=FALSE"
  Else
    db.CommitTrans
    MsgBox "Transaction Posted."
    'db2.Execute "DELETE * FROM [Bank Reconciliation]"
    grdDataGrid_HeadClick 0
    'Me.Requery
    'DoCmd GoToRecord A_FORM, "Bank Reconciliation", A_NEWREC
    Call LoadDetail
    'Call InitForm
  End If

  ShowStatus False
  
  'Me.RecordLocks = 1

  Exit Sub
  
RecordLocked:
  db.RollbackTrans
  Exit Sub

UnableToPost:
  db.RollbackTrans
  Exit Sub

cmdPost_Error:
  Call ErrorLog("Bank Reconciliation", "cmdPost_Click", Now, Err.Number, Err.Description, True, db)
  Resume Next

End Sub

Sub CalcRecTotals()

  'On Error GoTo CalcRecTotals_Error

  Dim in2$
  Dim Subtotal@
  Dim Total@
  Dim TotalOS@
  
  'Cleared Deposits
  in2$ = "('Transfer To','Cash Receipt','Deposit','Deposit Slip','Refund')"
  
  'Total@ = IIf(IsNull(SumRecord("[BANK RECD Amount]", "Bank Reconciliation Detail Credits", "[BANK RECD Type] in " & in2$ & " ")), 0, SumRecord("[BANK RECD Amount]", "Bank Reconciliation Detail Credits", "[BANK RECD Type] in " & in2$ & " "))
  'Forms![Bank Reconciliation].[Credits] = Total@
  
  Subtotal@ = IIf(IsNull(SumRecord("[BANK RECD Amount]", "[Bank Reconciliation Detail Credits]", db, "[BANK RECD Type] in " & in2$ & " AND [BANK RECD Cleared] = true")), 0, SumRecord("[BANK RECD Amount]", "[Bank Reconciliation Detail Credits]", db, "[BANK RECD Type] in " & in2$ & " AND [BANK RECD Cleared] = true"))
  txtLabels(12) = Subtotal@

  TotalOS@ = Total@ - Subtotal@
  txtLabels(6) = TotalOS@
  
  'Cleared Payments
  in2$ = "('Payment','Transfer From','Withdrawal')"
  
  'Total@ = IIf(IsNull(SumRecord("[BANK RECD Amount]", "Bank Reconciliation Detail Debits", "[BANK RECD Type] in " & in2$ & " ")), 0, SumRecord("[BANK RECD Amount]", "Bank Reconciliation Detail Debits", "[BANK RECD Type] in " & in2$ & " "))
  'Forms![Bank Reconciliation].[Debits] = Total@

  Subtotal@ = IIf(IsNull(SumRecord("[BANK RECD Amount]", "[Bank Reconciliation Detail Debits]", db, "[BANK RECD Type] in " & in2$ & " AND [BANK RECD Cleared] = true")), 0, SumRecord("[BANK RECD Amount]", "[Bank Reconciliation Detail Debits]", db, "[BANK RECD Type] in " & in2$ & " AND [BANK RECD Cleared] = true"))
  txtLabels(13) = Subtotal@

  TotalOS@ = Total@ - Subtotal@
  txtLabels(5) = TotalOS@
  
  
  'Cleared Balance
  'Forms![Bank Reconciliation]![Book Balance] = Forms![Bank Reconciliation]![Starting Balance] - Forms![Bank Reconciliation]![Debits] + Forms![Bank Reconciliation]![Credits] + Forms![Bank Reconciliation]![BANK REC Interest] - Forms![Bank Reconciliation]![BANK REC Service Charge]
  'MsgBox CDbl(txtLabels(0))
  'MsgBox CDbl(txtLabels(3))
  'MsgBox CDbl(txtLabels(2))
  'MsgBox CDbl(txtFields(6))
  'MsgBox CDbl(txtFields(8))
  txtLabels(4) = CDbl(txtLabels(0)) - CDbl(txtLabels(3)) + CDbl(txtLabels(2)) + CDbl(txtFields(6)) - CDbl(txtFields(8))
  
  'Forms![Bank Reconciliation]![Cleared Balance] = Forms![Bank Reconciliation]![Book Balance] + Forms![Bank Reconciliation]![Debits OS] - Forms![Bank Reconciliation]![Credits OS]
  txtLabels(7) = CDbl(txtLabels(4)) + CDbl(txtLabels(5)) - CDbl(txtLabels(6))
    
  'Forms![Bank Reconciliation].[Statement Balance] = Forms![Bank Reconciliation].[BANK REC Ending Balance]
  txtLabels(8) = txtFields(3)

  'Forms![Bank Reconciliation].[Difference] = Abs(Forms![Bank Reconciliation].[Cleared Balance] - Forms![Bank Reconciliation].[Statement Balance])
  txtLabels(9) = Abs(CDbl(txtLabels(7)) - CDbl(txtLabels(8)))

  'Forms![Bank Reconciliation]![Ending Book Balance] = Forms![Bank Reconciliation]![Starting Balance] + Forms![Bank Reconciliation]![Period Balance] + Forms![Bank Reconciliation]![BANK REC Interest] - Forms![Bank Reconciliation]![BANK REC Service Charge]
  txtLabels(10) = CDbl(txtLabels(0)) + CDbl(txtLabels(11)) + CDbl(txtFields(6)) - CDbl(txtFields(8))
  
  'Forms![Bank Reconciliation]![Starting Bank] = Forms![Bank Reconciliation]![Starting Balance] - ((Forms![Bank Reconciliation]![Credits] - Forms![Bank Reconciliation]![Debits]) - Forms![Bank Reconciliation]![Period Balance])
  txtLabels(1) = txtLabels(0) - ((CDbl(txtLabels(2)) - CDbl(txtLabels(3))) - CDbl(txtLabels(11)))
  
  'txtLabels(5) = FormatCurr(TotalOS@ )
  For Each Ctrl In Me.txtLabels
      Ctrl.Text = FormatCurr(Ctrl.Text)
  Next

  Exit Sub
CalcRecTotals_Error:
  Call ErrorLog("Bank Module", "CalcRecTotals", Now, Err.Number, Err.Description, True, db)
  Resume Next

End Sub


Private Sub InitForm()
    
  'On Error GoTo InitForm_Error
    
  'Load the first cash account
  'Dim rsGL As ADODB.Recordset
  'Set rsGL = New ADODB.Recordset
  'rsGL.Open "SELECT * FROM [GL Chart Of Accounts] where [GL COA Asset Type] = 'Cash'", db, adOpenKeyset, adLockOptimistic, adCmdText

  'On Error Resume Next

  'rsGL.MoveLast
  'rsGL.MoveFirst
  'If rsGL.RecordCount = 0 Then
  'Else
  '  rsGL.MoveFirst
  '  txtFields(1) = rsGL("GL COA Account No")
  'End If

  'On Error GoTo InitForm_Error

  'Load interest and service charge accounts
  Dim rsCompany As ADODB.Recordset
  Set rsCompany = New ADODB.Recordset
  rsCompany.Open "SELECT [SYS COM BANK Interest Earned Acct],[SYS COM BANK Service Charges Acct] FROM [SYS Company]", db, adOpenKeyset, adLockOptimistic, adCmdText
  rsCompany.MoveFirst
  
  ADOprimaryrs![BANK REC Interest Acct] = NZ(rsCompany("SYS COM BANK Interest Earned Acct"))
  ADOprimaryrs![BANK REC Service Acct] = NZ(rsCompany("SYS COM BANK Service Charges Acct"))
  ADOprimaryrs.Update
  
  rsCompany.Close
  Set rsCompany = Nothing
  
  Exit Sub
InitForm_Error:
  Call ErrorLog("Bank Reconciliation", "InitForm", Now, Err.Number, Err.Description, True, db)
  Resume Next

End Sub

Private Sub LoadDetail()

  'Fill Bank Reconciliation detail with all transactions for this bank

  'On Error GoTo LoadDetail_Error

  Dim BankID$
  BankID$ = NZ(ADOprimaryrs("BANK REC Bank Acct"))

  'Fill Starting Balance for this bank
  'xxx 12/4/96 7.2a
  Dim StartingBalance#
  Dim BookBalance#
  Dim PeriodBalance#
  Dim PeriodDebit#
  Dim PeriodCredit#
  
  Dim rsGL As ADODB.Recordset
  Set rsGL = New ADODB.Recordset
  rsGL.Open "SELECT [GL COA Account Balance] FROM [GL Chart Of Accounts] " & _
  "WHERE [GL COA Account No]='" & BankID$ & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
  'rsGL.Index = "PrimaryKey"
  'rsGL.Seek "=", BankID$
  rsGL.MoveFirst
  'rsGL.Find "[GL COA Account No]='" & BankID$ & "'"
  If rsGL.RecordCount = 0 Then
    MsgBox "Bank account not valid!", , "Error"
    StartingBalance# = 0
  Else
    StartingBalance# = rsGL("GL COA Account Balance")
    BookBalance# = rsGL("GL COA Account Balance")
    txtLabels(4) = BookBalance#
    PeriodBalance# = 0
    PeriodCredit# = 0
    PeriodDebit# = 0
  End If
  
  rsGL.Close
  Set rsGL = Nothing
  
' Load Credits
  'Debug.Print "load credits"
  db.Execute "DELETE * FROM [Bank Reconciliation Detail Credits]"
  
  'Dim rsDetail As ADODB.Recordset
  Dim rsQuery As ADODB.Recordset

  'xxx 4/14/97 7.3 Changed cleared to reconciled
  'Set rsQuery = db.OpenRecordset("SELECT * FROM [qryBankTransactions] WHERE [BANK ID] = '" & BankID$ & "' AND [Cleared] = False AND [Type] in ('Deposit','Cash Receipt','Transfer To')")
  Set rsQuery = New ADODB.Recordset
  'rsQuery.Open "SELECT * FROM [qryBankTransactions] WHERE [BANK ID] = '" & BankID$ & "' AND [Type] in ('Deposit','Deposit Slip','Cash Receipt','Transfer To','Refund')", db, adOpenKeyset, adLockOptimistic, adCmdText
  'rsQuery.Open "SELECT [Date],[Cleared],[Doc #],[Description],[Type],[Amount] " & _
  '"FROM [qryBankTransactions] WHERE [BANK ID] = '" & BankID$ & "' AND [Type] in ('Payment','Payroll','Withdrawal','Transfer From')", db, adOpenKeyset, adLockOptimistic, adCmdText
  rsQuery.Open "SELECT [Date],[Cleared],[Doc #],[Description],[Type],[Amount] " & _
  "FROM [qryBankTransactions] WHERE [BANK ID] = '" & BankID$ & "' AND [Type] in ('Deposit','Deposit Slip','Cash Receipt','Transfer To','Refund')", db, adOpenKeyset, adLockOptimistic, adCmdText
  
  'Set rsDetail = New ADODB.Recordset
  'rsDetail.Open "[Bank Reconciliation Detail Credits]", db, adOpenKeyset, adLockOptimistic, adCmdTable

  'On Error Resume Next
  'If Err <> 0 Then
  If rsQuery.RecordCount = 0 Then
    MsgBox "No open credit transactions for this bank!"
    GoTo processdebits
    Else
  End If
  
  rsQuery.MoveFirst

  'On Error GoTo LoadDetail_Error

  Do While Not rsQuery.EOF
    If rsQuery("Date") <= ADOprimaryrs![BANK REC Cutoff Date] And rsQuery("Cleared") = False Then
        'rsDetail.AddNew
        
            SQLstatement = "INSERT INTO [Bank Reconciliation Detail Credits]"
            SQLstatement = SQLstatement & " ([BANK RECD Doc #],[BANK RECD Cleared]," & _
            "[BANK RECD Date],[BANK RECD Description],[BANK RECD Type],[BANK RECD Amount])"
            
            SQLstatement = SQLstatement & " VALUES ('" & rsQuery("Doc #") & "',False,#" & _
            rsQuery("Date") & "#,'" & rsQuery("Description") & "','" & rsQuery("Type") & "'," & _
            rsQuery("Amount") & ")"
            db.Execute SQLstatement
          
          'rsDetail("BANK RECD Doc #") = rsQuery("Doc #") & ""
          'rsDetail("BANK RECD Cleared") = False
          'rsDetail("BANK RECD Date") = rsQuery("Date")
          'rsDetail("BANK RECD Description") = rsQuery("Description") & ""
          'rsDetail("BANK RECD Type") = rsQuery("Type") & ""
          'rsDetail("BANK RECD Amount") = rsQuery("Amount")
          
          If rsQuery("Date") >= ADOprimaryrs![BANK REC Start Date] Then
            Select Case rsQuery("Type")
            Case "Cash Receipt", "Transfer To", "Deposit", "Deposit Slip", "Refund"
              StartingBalance# = StartingBalance# - rsQuery("Amount")
              PeriodBalance# = PeriodBalance# + rsQuery("Amount")
              PeriodCredit# = PeriodCredit# + rsQuery("Amount")
              'Debug.Print "credit", PeriodBalance#, rsQuery("Amount")
            Case "Payment", "Transfer From", "Withdrawal"
              StartingBalance# = StartingBalance# + rsQuery("Amount")
              PeriodBalance# = PeriodBalance# - rsQuery("Amount")
              PeriodCredit# = PeriodCredit# - rsQuery("Amount")
              'Debug.Print "credit", PeriodBalance#, rsQuery("Amount")
            End Select
          Else
          End If
          'rsDetail.Update
    Else
            If rsQuery("Date") >= ADOprimaryrs![BANK REC Start Date] Then
                Select Case rsQuery("Type")
                Case "Cash Receipt", "Transfer To", "Deposit", "Deposit Slip", "Refund"
                  StartingBalance# = StartingBalance# - rsQuery("Amount")
                Case "Payment", "Transfer From", "Withdrawal"
                  StartingBalance# = StartingBalance# + rsQuery("Amount")
                End Select
                Else
            End If
    End If
    rsQuery.MoveNext
  Loop
  rsQuery.Close
  Set rsQuery = Nothing
  
  'rsDetail.Close
  'Set rsDetail = Nothing
  
processdebits:
  
' Load Debits
'Debug.Print "load debits"
  db.Execute "DELETE * FROM [Bank Reconciliation Detail Debits]"
  
  'Dim rsDetail As ADODB.Recordset
  'Dim rsQuery As ADODB.Recordset

  'xxx 4/14/97 7.3 Changed cleared to reconciled
  'Set rsQuery = db.OpenRecordset("SELECT * FROM [qryBankTransactions] WHERE [BANK ID] = '" & BankID$ & "' AND [Cleared] = False AND [Type] in ('Payment','Withdrawal','Transfer From')")
  Set rsQuery = New ADODB.Recordset
  rsQuery.Open "SELECT [Date],[Cleared],[Doc #],[Description],[Type],[Amount] " & _
  "FROM [qryBankTransactions] WHERE [BANK ID] = '" & BankID$ & "' AND [Type] in ('Payment','Payroll','Withdrawal','Transfer From')", db, adOpenKeyset, adLockOptimistic, adCmdText
  
  'Set rsDetail = New ADODB.Recordset
  'rsDetail.Open "[Bank Reconciliation Detail Debits]", db, adOpenKeyset, adLockOptimistic, adCmdTable

  'On Error Resume Next
  If rsQuery.RecordCount = 0 Then
    MsgBox "No open debit transactions for this bank!"
    GoTo skipdebits
    Else
  End If
  rsQuery.MoveFirst

  'On Error GoTo LoadDetail_Error

    Do While Not rsQuery.EOF
    If rsQuery("Date") <= ADOprimaryrs![BANK REC Cutoff Date] And rsQuery("Cleared") = False Then
        'rsDetail.AddNew
        
            SQLstatement = "INSERT INTO [Bank Reconciliation Detail Debits]"
            SQLstatement = SQLstatement & " ([BANK RECD Doc #],[BANK RECD Cleared]," & _
            "[BANK RECD Date],[BANK RECD Description],[BANK RECD Type],[BANK RECD Amount])"
            
            SQLstatement = SQLstatement & " VALUES ('" & rsQuery("Doc #") & "',False,#" & _
            rsQuery("Date") & "#,'" & rsQuery("Description") & "','" & rsQuery("Type") & "'," & _
            rsQuery("Amount") & ")"
            db.Execute SQLstatement
            
          'rsDetail("BANK RECD Doc #") = rsQuery("Doc #") & ""
          'rsDetail("BANK RECD Cleared") = False
          'rsDetail("BANK RECD Date") = rsQuery("Date")
          'rsDetail("BANK RECD Description") = rsQuery("Description") & ""
          'rsDetail("BANK RECD Type") = rsQuery("Type") & ""
          'rsDetail("BANK RECD Amount") = rsQuery("Amount")
          If rsQuery("Date") >= ADOprimaryrs![BANK REC Start Date] Then
            Select Case rsQuery("Type")
            Case "Cash Receipt", "Transfer To", "Deposit", "Deposit Slip", "Refund"
              StartingBalance# = StartingBalance# - rsQuery("Amount")
              PeriodBalance# = PeriodBalance# + rsQuery("Amount")
              PeriodDebit# = PeriodDebit# + rsQuery("Amount")
              'Debug.Print "debit", PeriodBalance#, rsQuery("Amount")
            Case "Payment", "Transfer From", "Withdrawal"
              StartingBalance# = StartingBalance# + rsQuery("Amount")
              PeriodBalance# = PeriodBalance# - rsQuery("Amount")
              PeriodDebit# = PeriodDebit# - rsQuery("Amount")
              'Debug.Print "debit", PeriodBalance#, rsQuery("Amount")
            End Select
            Else
          End If
          'rsDetail.Update
    Else
            If rsQuery("Date") >= ADOprimaryrs![BANK REC Start Date] Then
                Select Case rsQuery("Type")
                Case "Cash Receipt", "Transfer To", "Deposit", "Deposit Slip", "Refund"
                  StartingBalance# = StartingBalance# - rsQuery("Amount")
                Case "Payment", "Transfer From", "Withdrawal"
                  StartingBalance# = StartingBalance# + rsQuery("Amount")
                End Select
                Else
            End If
    End If
    rsQuery.MoveNext
  Loop
skipdebits:
  
  rsQuery.Close
  Set rsQuery = Nothing
  
  'rsDetail.Close
  'Set rsDetail = Nothing
  
  txtLabels(0) = StartingBalance#
  txtLabels(11) = PeriodBalance#
  txtLabels(2) = PeriodCredit#
  txtLabels(3) = PeriodDebit#
  'Me.Refresh
  
  'Me![Bank Reconciliation Detail Debits].Requery
  'Me![Bank Reconciliation Detail Credits].Requery
  
  Exit Sub
LoadDetail_Error:
  Call ErrorLog("Bank Reconciliation", "LoadDetail", Now, Err.Number, Err.Description, True, db)
  Resume Next

End Sub

Private Sub cmdRecalc_Click()

  'On Error GoTo cmdRecalc_Click_Error
  ShowStatus True
  Call LoadDetail
  Call CalcRecTotals
  ShowStatus False
  MsgBox "Process Done."

  Exit Sub
cmdRecalc_Click_Error:
  Call ErrorLog("Bank Reconciliation", "cmdRecalc_Click", Now, Err.Number, Err.Description, True, db)
  Resume Next

End Sub

Private Sub cmdSearch_Click()
On Error GoTo NOTFOUND
    If txtFields(0) = "" Then Exit Sub
    grdDataGrid.SetFocus
    adoPrimaryRS2.MoveFirst
    
    If WhichField = "" Then
        WhichField = "BANK RECD Doc #"
    End If
    If adoPrimaryRS2("" & WhichField & "").Type = 202 Then
        adoPrimaryRS2.Find "[" & WhichField & "]='" & txtFields(0) & "'"
    Else
        adoPrimaryRS2.Find "[" & WhichField & "]=" & txtFields(0)
    End If
    If adoPrimaryRS2.EOF Then
NOTFOUND:
        MsgBox lblLabels(24) & " " & txtFields(0) & " is not existed.", vbInformation, "Information"
    End If
    SendKeys ("{ENTER}")
End Sub


Private Sub Form_Load()
'On Error GoTo FormErr
  Set db = New ADODB.Connection
  db.CursorLocation = adUseClient
  db.Open gblADOProvider

  LoadEverything
  
  'Set adoPrimaryRS = New ADODB.Recordset
  'Set adoPrimaryRS2 = New ADODB.Recordset
  
  GetTextColor Me
  mbDataChanged = False
Exit Sub
FormErr:
  MsgBox Err.Description
  ShowStatus False
End Sub

Private Sub LoadEverything()
ShowStatus True
  FormLoadData
  Call InitForm
  Call LoadDetail
  
  ADOprimaryrs.Requery
  Dim Ctrl As Control
  For Each Ctrl In Me.txtFields
      Set Ctrl.DataSource = ADOprimaryrs
  Next
  
  Call CalcRecTotals
  
  'Call ReconcileDetail_AfterUpdate
  
  'Set grdDataGrid.DataSource = adoPrimaryRS2
  If optType(1).Value = True Then
    optType(0).Value = True
  Else
    optType(1).Value = True
  End If
  'optType_Click 0
  ShowStatus False
End Sub

Private Sub Form_Resize()
  If fMainForm.WindowState = 1 Then Exit Sub
  If Me.WindowState = 0 Then
  ElseIf Me.WindowState = 2 Then
    GoTo SkipResize
  Else
    Exit Sub
  End If
  
  Me.Width = 9990
  Me.Height = 9075
  
SkipResize:
  frPrimary.Left = (Me.ScaleWidth - frPrimary.Width) / 2
  Label1(1).Left = frPrimary.Left
  Label1(1).Width = frPrimary.Width
  frPrimary.Top = (Me.ScaleHeight - frPrimary.Height) / 2 + 230

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If mbAddNewFlag Then Exit Sub

  Select Case KeyCode
    Case vbKeyEscape
      cmdClose_Click
    Case vbKeyEnd
      cmdLast_Click
    Case vbKeyHome
      cmdFirst_Click
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
  ShowStatus True
      If ADOprimaryrs.RecordCount > 0 Then
        If ADOprimaryrs.EditMode <> 0 Then
          ADOprimaryrs.CancelUpdate
        End If
      End If
      ADOprimaryrs.Close
      Set ADOprimaryrs = Nothing
      db.Close
      Set db = Nothing
  ShowStatus False
Exit Sub
FormErr:
  MsgBox Err.Description
  ShowStatus False
End Sub

Private Sub adoPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This will display the current record position for this recordset
  lblStatus.Caption = "Record: " & CStr(ADOprimaryrs.AbsolutePosition) & " of " & CStr(ADOprimaryrs.RecordCount)
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
  With ADOprimaryrs
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
    .AddNew
    lblStatus.Caption = "Add record"
    mbAddNewFlag = True
    SetButtons True
  End With

  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
     If Not DataDelete(ADOprimaryrs, Me, True) Then
     End If
End Sub

Private Sub cmdRefresh_Click()
  'This is only needed for multi user apps
  'On Error GoTo RefreshErr
  Set grdDataGrid.DataSource = Nothing
  ADOprimaryrs.Requery
  Set grdDataGrid.DataSource = ADOprimaryrs("ChildCMD").UnderlyingValue
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub

Private Sub cmdEdit_Click()
  'On Error GoTo EditErr

  lblStatus.Caption = "Edit record"
  mbEditFlag = True
  SetButtons True
  Exit Sub

EditErr:
  MsgBox Err.Description
End Sub
Private Sub cmdCancel_Click()
  'On Error Resume Next

  SetButtons True
  mbEditFlag = False
  mbAddNewFlag = False
  ADOprimaryrs.CancelUpdate
  If mvBookMark > 0 Then
    ADOprimaryrs.Bookmark = mvBookMark
  Else
    ADOprimaryrs.MoveFirst
  End If
  mbDataChanged = False

End Sub

Private Sub cmdUpdate_Click()
'Dim FlagStatus As Boolean
    
  'FlagStatus = False

  Call UpdateButton(ADOprimaryrs, mbAddNewFlag)
  
  'mbEditFlag = Not FlagStatus
  
  'SetButtons FlagStatus
  'mbDataChanged = Not FlagStatus
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdFirst_Click()
  'On Error GoTo GoFirstError

  ADOprimaryrs.MoveFirst
  mbDataChanged = False

  Exit Sub

GoFirstError:
  MsgBox Err.Description
End Sub

Private Sub cmdLast_Click()
  'On Error GoTo GoLastError

  ADOprimaryrs.MoveLast
  mbDataChanged = False

  Exit Sub

GoLastError:
  MsgBox Err.Description
End Sub

Private Sub cmdNext_Click()
  'On Error GoTo GoNextError

  If Not ADOprimaryrs.EOF Then ADOprimaryrs.MoveNext
  If ADOprimaryrs.EOF And ADOprimaryrs.RecordCount > 0 Then
    Beep
     'moved off the end so go back
    ADOprimaryrs.MoveLast
  End If
  'show the current record
  mbDataChanged = False

  Exit Sub
GoNextError:
  MsgBox Err.Description
End Sub

Private Sub cmdPrevious_Click()
  'On Error GoTo GoPrevError

  If Not ADOprimaryrs.BOF Then ADOprimaryrs.MovePrevious
  If ADOprimaryrs.BOF And ADOprimaryrs.RecordCount > 0 Then
    Beep
    'moved off the end so go back
    ADOprimaryrs.MoveFirst
  End If
  'show the current record
  mbDataChanged = False

  Exit Sub

GoPrevError:
  MsgBox Err.Description
End Sub

Private Sub SetButtons(bVal As Boolean)
  cmdAdd.Visible = bVal
  cmdUpdate.Visible = bVal
  cmdCancel.Visible = Not bVal
  cmdDelete.Visible = bVal
  cmdClose.Visible = bVal
  cmdRefresh.Visible = bVal
  cmdNext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdLast.Enabled = bVal
  cmdPrevious.Enabled = bVal
End Sub

Private Sub grdDataGrid_ButtonClick(ByVal ColIndex As Integer)
    If grdDataGrid.Row = -1 Or grdDataGrid.Columns(0) = "" Then Exit Sub
         SendKeys ("{ENTER}")
   If grdDataGrid.Columns(1).Text = "No" Then
      grdDataGrid.Columns(1).Text = "Yes"
   Else
      grdDataGrid.Columns(1).Text = "No"
   End If
         SendKeys ("{ENTER}")
         SendKeys ("{down}")
         SendKeys ("{up}")
         adoPrimaryRS2.Update
         CalcTotals
         CalcRecTotals

End Sub

Private Sub CalcTotals()
Dim TempCount As Currency
Dim TempSQlOpen As ADODB.Recordset
Dim TemptxtLabels As TextBox
Dim TemptxtFields As TextBox

  Set TempSQlOpen = New ADODB.Recordset
  If optType(0).Value = True Then
     Set TemptxtLabels = txtLabels(12)
     Set TemptxtFields = txtLabels(6)
     SelectedTable = "[BANK Reconciliation Detail Credits]"
  Else
     Set TemptxtLabels = txtLabels(13)
     Set TemptxtFields = txtLabels(5)
     SelectedTable = "[BANK Reconciliation Detail Debits]"
  End If
  
  TempSQlOpen.Open "SELECT [BANK RECD Doc #],[BANK RECD Cleared],[BANK RECD Date],[BANK RECD Description],[BANK RECD Type],[BANK RECD Amount] FROM " & SelectedTable, db, adOpenKeyset, adLockOptimistic, adCmdText
  With TempSQlOpen
    If .RecordCount = 0 Then
        TemptxtLabels = "$0.00"
        TemptxtFields = "$0.00"
      Exit Sub
    End If
    .MoveFirst
    TempCount = 0
    Do While Not .EOF
      If ![BANK RECD Cleared] = True Then
        TempCount = TempCount + ![BANK RECD Amount]
      End If
        .MoveNext
    Loop
  End With
        TemptxtLabels = FormatCurr(TempCount)
        TemptxtFields = TemptxtLabels
End Sub

Private Sub grdDataGrid_HeadClick(ByVal ColIndex As Integer)
If adoPrimaryRS2.RecordCount = 0 Then Exit Sub 'Label1
    lblLabels(24) = grdDataGrid.Columns(ColIndex).Caption
    WhichField = grdDataGrid.Columns(ColIndex).DataField
    adoPrimaryRS2.Close
    Set adoPrimaryRS2 = Nothing
    Set adoPrimaryRS2 = New ADODB.Recordset
    adoPrimaryRS2.Open TempStr & " ORDER BY [" & WhichField & "]", db, adOpenKeyset, adLockOptimistic, adCmdText
    Set grdDataGrid.DataSource = adoPrimaryRS2
End Sub


Private Sub Label1_Click(Index As Integer)
Dim TotalChar, i, j As Integer
Dim storedString As String
Dim SQLstatement As String
Dim grdDataGridHeader As String
 
 grdDataGridHeader = GridHeader

  Dim grdWidth As Integer
  grdWidth = 0
  'Change The datagrid columnheader name
TotalChar = Len(grdDataGridHeader)
    For i = 1 To TotalChar

         If Mid(grdDataGridHeader, i, 2) = "//" Then
            grdDataGridHeader = Right(grdDataGridHeader, TotalChar - i - 1)
            TotalChar = Len(grdDataGridHeader)
            grdAllLookup.Columns(j).Caption = storedString
            
            If Len(Trim(ADOprimaryrs("" & grdAllLookup.Columns(j).DataField & ""))) < 8 Then
                grdAllLookup.Columns(j).Width = 800
            ElseIf Len(Trim(ADOprimaryrs("" & grdAllLookup.Columns(j).DataField & ""))) < 30 Then
                grdAllLookup.Columns(j).Width = 2000
            Else
                grdAllLookup.Columns(j).Width = 2000
            End If
            grdWidth = grdWidth + grdAllLookup.Columns(j).Width
            
            j = j + 1
            storedString = ""
            i = 1
            If TotalChar = 0 Then Exit For
         End If
         storedString = Left(grdDataGridHeader, i)
         
    Next
    grdAllLookup.Columns(j).Caption = storedString
    If Len(Trim(ADOprimaryrs("" & grdAllLookup.Columns(j).DataField & ""))) < 8 Then
        grdAllLookup.Columns(j).Width = 800
    ElseIf Len(Trim(ADOprimaryrs("" & grdAllLookup.Columns(j).DataField & ""))) < 30 Then
        grdAllLookup.Columns(j).Width = 2000
    Else
        grdAllLookup.Columns(j).Width = 2000
    End If
    grdWidth = grdWidth + grdAllLookup.Columns(j).Width
    
  'hide unneeded columns
  For i = j + 1 To grdAllLookup.Columns.count - 1
    grdAllLookup.Columns(i).Visible = False
  Next
  End Sub

Private Sub optType_Click(Index As Integer)
Dim SelectedTable As String
Select Case Index
Case 0
     SelectedTable = "[BANK Reconciliation Detail Credits]"
Case 1
     SelectedTable = "[BANK Reconciliation Detail Debits]"
End Select
   Set adoPrimaryRS2 = New ADODB.Recordset
   TempStr = "SELECT [BANK RECD Doc #],[BANK RECD Cleared],[BANK RECD Date],[BANK RECD Description],[BANK RECD Type],[BANK RECD Amount]FROM " & SelectedTable
   adoPrimaryRS2.Open TempStr, db, adOpenKeyset, adLockOptimistic, adCmdText
   
   Set grdDataGrid.DataSource = adoPrimaryRS2
   grdDataGrid.Columns(1).Button = True
End Sub

Private Sub txtfields_KeyPress(Index As Integer, KeyAscii As Integer)
Dim keyResponse As Boolean
Select Case Index
Case 0
    If KeyCode = 13 Then cmdSearch_Click
Case 3
    keyResponse = CtrlValidate(KeyAscii, "-.0123456789")
    If keyResponse = False Then
       KeyAscii = 0
    End If
Case 6, 8
    keyResponse = CtrlValidate(KeyAscii, ".0123456789")
    If keyResponse = False Then
       KeyAscii = 0
    End If
End Select
End Sub

Private Sub txtFields_LostFocus(Index As Integer)
Select Case Index
Case 3
    txtFields(3) = FormatCurr(txtFields(3))
    txtLabels(8) = txtFields(3)
    CalcRecTotals
Case 6, 8
    txtFields(Index) = FormatCurr(txtFields(Index))
    CalcRecTotals
End Select
End Sub
