VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_AP_Cash_Payments 
   Caption         =   "Cash Receipts"
   ClientHeight    =   7215
   ClientLeft      =   2070
   ClientTop       =   2715
   ClientWidth     =   10215
   Icon            =   "frm_AP_Cash_Payments.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7215
   ScaleWidth      =   10215
   Begin VB.Frame frPrimary 
      Height          =   6735
      Left            =   0
      TabIndex        =   15
      Top             =   480
      Width           =   10215
      Begin VB.TextBox txtFields 
         Alignment       =   2  'Center
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   63
         Top             =   4200
         Width           =   2655
      End
      Begin VB.TextBox txtFields 
         DataField       =   "AR PAY Customer No"
         Height          =   285
         Index           =   0
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox txtFields 
         DataField       =   "AR PAY Bank Account"
         Height          =   285
         Index           =   1
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   2640
         Width           =   1455
      End
      Begin VB.TextBox txtFields 
         DataField       =   "AR PAY ID"
         Height          =   285
         Index           =   2
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   1920
         Width           =   1815
      End
      Begin VB.TextBox txtFields 
         DataField       =   "AR PAY Transaction Date"
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox txtFields 
         DataField       =   "AR PAY Amount"
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
         Left            =   7320
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   3240
         Width           =   1575
      End
      Begin VB.TextBox txtFields 
         DataField       =   "AR PAY Check No"
         Height          =   285
         Index           =   5
         Left            =   1560
         TabIndex        =   35
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox txtFields 
         DataField       =   "AR PAY Type"
         Height          =   285
         Index           =   6
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   2280
         Width           =   1815
      End
      Begin VB.TextBox txtFields 
         DataField       =   "AR PAY Status"
         Height          =   285
         Index           =   11
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   3600
         Width           =   2655
      End
      Begin VB.TextBox txtFields 
         DataField       =   "AR PAY UnApplied Amount"
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
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   3240
         Width           =   1575
      End
      Begin VB.TextBox txtFields 
         DataField       =   "AR PAY Notes"
         Height          =   885
         Index           =   14
         Left            =   5040
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   31
         Top             =   3600
         Width           =   3855
      End
      Begin VB.TextBox TotalApplied 
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
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   3240
         Width           =   1575
      End
      Begin VB.TextBox TotalDue 
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   3240
         Width           =   1575
      End
      Begin VB.CommandButton cmdCashCustomer 
         Height          =   285
         Left            =   3000
         Picture         =   "frm_AP_Cash_Payments.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   840
         Width           =   375
      End
      Begin VB.CommandButton cmdCheckNo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3000
         Picture         =   "frm_AP_Cash_Payments.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   1560
         Width           =   375
      End
      Begin VB.CommandButton cmdNSF 
         Caption         =   "&NSF"
         Enabled         =   0   'False
         Height          =   975
         Left            =   9000
         Picture         =   "frm_AP_Cash_Payments.frx":059E
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CommandButton cmdApply 
         Caption         =   "&Apply"
         Enabled         =   0   'False
         Height          =   975
         Left            =   9000
         Picture         =   "frm_AP_Cash_Payments.frx":08A8
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   2400
         Width           =   1095
      End
      Begin VB.CommandButton cmdbankAccount 
         Height          =   285
         Left            =   3000
         Picture         =   "frm_AP_Cash_Payments.frx":0BB2
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   2640
         Width           =   375
      End
      Begin VB.Frame Frame1 
         Enabled         =   0   'False
         Height          =   1575
         Left            =   3480
         TabIndex        =   16
         Top             =   1320
         Width           =   5415
         Begin VB.CheckBox chkFields 
            Alignment       =   1  'Right Justify
            Caption         =   "NSF:"
            DataField       =   "AR PAY NSF"
            Height          =   255
            Index           =   13
            Left            =   3600
            TabIndex        =   21
            Top             =   720
            Width           =   1455
         End
         Begin VB.CheckBox chkFields 
            Alignment       =   1  'Right Justify
            Caption         =   "Reconciled:"
            DataField       =   "AR PAY Reconciled"
            Height          =   255
            Index           =   10
            Left            =   360
            TabIndex        =   20
            Top             =   720
            Width           =   1455
         End
         Begin VB.CheckBox chkFields 
            Alignment       =   1  'Right Justify
            Caption         =   "Posted:"
            DataField       =   "AR PAY Posted YN"
            Height          =   255
            Index           =   9
            Left            =   360
            TabIndex        =   19
            Top             =   360
            Width           =   1455
         End
         Begin VB.CheckBox chkFields 
            Alignment       =   1  'Right Justify
            Caption         =   "Cleared:"
            DataField       =   "AR PAY Cleared"
            Height          =   255
            Index           =   7
            Left            =   360
            TabIndex        =   18
            Top             =   1080
            Width           =   1455
         End
         Begin VB.CheckBox chkFields 
            Alignment       =   1  'Right Justify
            Caption         =   "Deposited:"
            DataField       =   "AR PAY Deposited YN"
            Height          =   255
            Index           =   8
            Left            =   3600
            TabIndex        =   17
            Top             =   360
            Width           =   1455
         End
      End
      Begin MSDataGridLib.DataGrid grdDataGrid 
         Height          =   2025
         Left            =   120
         TabIndex        =   30
         Top             =   4560
         Width           =   8760
         _ExtentX        =   15452
         _ExtentY        =   3572
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
         ColumnCount     =   8
         BeginProperty Column00 
            DataField       =   "Reference"
            Caption         =   "Reference"
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
            DataField       =   "Date"
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
         BeginProperty Column02 
            DataField       =   "Original Amount"
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
            DataField       =   "Amount Paid"
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
         BeginProperty Column04 
            DataField       =   "Discount"
            Caption         =   "Discount"
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
            DataField       =   "Write Off"
            Caption         =   "Write Off"
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
            DataField       =   "Applied Amount"
            Caption         =   "Applied"
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
         BeginProperty Column07 
            DataField       =   "Balance"
            Caption         =   "Balance"
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
               ColumnWidth     =   1094.74
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1094.74
            EndProperty
            BeginProperty Column02 
               Alignment       =   1
               ColumnWidth     =   1094.74
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               ColumnWidth     =   1094.74
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column06 
               Alignment       =   1
               ColumnWidth     =   1094.74
            EndProperty
            BeginProperty Column07 
               Alignment       =   1
               ColumnWidth     =   1094.74
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   975
         Left            =   9000
         Picture         =   "frm_AP_Cash_Payments.frx":0CFC
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Frame frNSF 
         Height          =   2055
         Left            =   9000
         TabIndex        =   55
         Top             =   1320
         Visible         =   0   'False
         Width           =   1095
         Begin VB.Label Label3 
            Caption         =   "und"
            Height          =   255
            Index           =   2
            Left            =   480
            TabIndex        =   61
            Top             =   1680
            Width           =   375
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "ufficient"
            Height          =   255
            Index           =   1
            Left            =   480
            TabIndex        =   60
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label Label3 
            Caption         =   "on"
            Height          =   255
            Index           =   0
            Left            =   600
            TabIndex        =   59
            Top             =   480
            Width           =   375
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "F"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   36
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   825
            Index           =   2
            Left            =   120
            TabIndex        =   58
            Top             =   1200
            Width           =   405
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "S"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   36
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   825
            Index           =   1
            Left            =   120
            TabIndex        =   57
            Top             =   600
            Width           =   435
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "N"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   36
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   825
            Index           =   0
            Left            =   120
            TabIndex        =   56
            Top             =   0
            Width           =   525
         End
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Prin&t"
         Height          =   975
         Left            =   9000
         Picture         =   "frm_AP_Cash_Payments.frx":1006
         Style           =   1  'Graphical
         TabIndex        =   65
         ToolTipText     =   "Print Customer Customer Statement"
         Top             =   5640
         Width           =   1095
      End
      Begin VB.CommandButton cmdPost 
         Caption         =   "&Post"
         Enabled         =   0   'False
         Height          =   975
         Left            =   9000
         Picture         =   "frm_AP_Cash_Payments.frx":1310
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   4560
         Width           =   1095
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Account Balances"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   1560
         TabIndex        =   64
         Top             =   3960
         Width           =   2655
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Customer ID:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   53
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Bank Account:"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   52
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Payment ID:"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   51
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Transaction Date:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   50
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Amount:"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   7320
         TabIndex        =   49
         Top             =   3000
         Width           =   1575
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Check No:"
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   48
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Payment Type:"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   47
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Status:"
         Height          =   255
         Index           =   11
         Left            =   600
         TabIndex        =   46
         Top             =   3600
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "UnApplied Amount:"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   12
         Left            =   5400
         TabIndex        =   45
         Top             =   3000
         Width           =   1575
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Notes:"
         Height          =   255
         Index           =   14
         Left            =   4080
         TabIndex        =   44
         Top             =   3600
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Applied Amount:"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   15
         Left            =   3480
         TabIndex        =   43
         Top             =   3000
         Width           =   1575
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total Due:"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   16
         Left            =   1560
         TabIndex        =   42
         Top             =   3000
         Width           =   1575
      End
      Begin VB.Label lblcashReceiptsTrue 
         Alignment       =   2  'Center
         Caption         =   "Customer name"
         BeginProperty Font 
            Name            =   "Vibrocentric"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   120
         TabIndex        =   41
         Top             =   240
         Width           =   9975
      End
   End
   Begin VB.TextBox txtcustNo 
      Height          =   285
      Left            =   9240
      TabIndex        =   14
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox txtcheckNo 
      Height          =   285
      Left            =   9240
      TabIndex        =   13
      Top             =   2760
      Width           =   855
   End
   Begin VB.TextBox txtID 
      Height          =   285
      Left            =   9240
      TabIndex        =   12
      Top             =   2040
      Width           =   855
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   10215
      TabIndex        =   6
      Top             =   6615
      Visible         =   0   'False
      Width           =   10215
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   300
         Left            =   4320
         TabIndex        =   10
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   300
         Left            =   3240
         TabIndex        =   9
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   300
         Left            =   2160
         TabIndex        =   8
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   300
         Left            =   1080
         TabIndex        =   11
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   300
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.PictureBox picStatBox 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   10215
      TabIndex        =   0
      Top             =   6915
      Visible         =   0   'False
      Width           =   10215
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Picture         =   "frm_AP_Cash_Payments.frx":1752
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Picture         =   "frm_AP_Cash_Payments.frx":1A94
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Picture         =   "frm_AP_Cash_Payments.frx":1DD6
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Picture         =   "frm_AP_Cash_Payments.frx":2118
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   5
         Top             =   0
         Width           =   3360
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cash Receipt"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   2
      Left            =   20
      TabIndex        =   54
      Top             =   120
      Width           =   10185
   End
End
Attribute VB_Name = "frm_AP_Cash_Payments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ADOprimaryrs As ADODB.Recordset
Attribute ADOprimaryrs.VB_VarHelpID = -1
Dim WithEvents adoPrimaryRS2 As ADODB.Recordset
Attribute adoPrimaryRS2.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean

Dim PrevcheckNo As String

Dim db As ADODB.Connection



Private Sub CallRebuildTable()
If txtfields(0).Text = "" Then Exit Sub
Screen.MousePointer = vbHourglass

    'If ValidateData = True Then Exit Sub  '<<<---- do we need this
    'CheckAllData (True)
    CheckAllData
    RebuildTable
    grdDataGridSource "select *  from [Cash Receipts Work] Order by [Reference]"
Screen.MousePointer = vbNormal
End Sub

Private Sub LockControl()
Dim i As Integer
Dim EmptyTxt As Boolean

If cmdApply.Caption = "&Save" Then
    cmdApply.Enabled = True
    cmdDel.Enabled = True
    Exit Sub
End If

EmptyTxt = False
For i = txtfields.LBound To txtfields.UBound
    Select Case i
    Case 0, 1, 2, 4, 5
        If txtfields(i) = "" Then
            EmptyTxt = True
        End If
    End Select
Next
If EmptyTxt = False Then
    cmdDel.Enabled = True
    cmdApply.Enabled = True
    cmdNSF.Enabled = False
    cmdPost.Enabled = True
    frNSF.Visible = False
    
    For i = chkFields.LBound To chkFields.UBound
        Select Case i
        Case 7, 10, 13
          If chkFields(i).Value = 1 Then
            cmdDel.Enabled = False
            cmdApply.Enabled = False
            cmdNSF.Enabled = False
            cmdPost.Enabled = False
            EmptyTxt = True
          End If
        Case 9
          If chkFields(i).Value = 1 Then
            cmdDel.Enabled = False
            cmdApply.Enabled = False
            cmdNSF.Enabled = True
            cmdPost.Enabled = False
            EmptyTxt = True
          End If
        End Select
    Next
    If EmptyTxt = False Then
        cmdApply.Enabled = True
        cmdDel.Enabled = True
    End If
Else
    cmdDel.Enabled = False
    cmdApply.Enabled = False
    cmdNSF.Enabled = False
    cmdPost.Enabled = False
    frNSF.Visible = False
End If
End Sub

Private Sub PostData()

'  On Error GoTo PostData_Error
  
  Dim rsWork As ADODB.Recordset
  'Dim rsCross As ADODB.Recordset
  Dim rsSales As ADODB.Recordset
  Dim SaleID$
  Dim SQLstatement As String


  If txtfields(2) = "" Then Exit Sub

  Set rsWork = New Recordset
  rsWork.Open "SELECT * FROM [Cash Receipts Work]", db, adOpenForwardOnly, adLockOptimistic
  
  'Set rsCross = New Recordset
  'rsCross.Open "SELECT * FROM [AR Payment Invoice Cross Reference]", db, adOpenStatic, adLockOptimistic
  
  Set rsSales = New Recordset
  rsSales.Open "SELECT [AR SALE Total],[AR SALE Amount Paid],[AR SALE Balance Due],[AR SALE Document #],[AR SALE Ext Document #] FROM [AR Sales]", db, adOpenForwardOnly, adLockOptimistic

  'rsSales.Index = "Ext Document #"

  'Scratch detail
  db.Execute "DELETE * FROM [AR Payment Invoice Cross Reference] where [AR CROSS Payment ID] = " & txtfields(2).Text, , adCmdText
  
  'Save detail information
  'On Error Resume Next
  rsWork.MoveFirst
  'If Err = 0 Then
    'On Error GoTo PostData_Error
    Do While Not rsWork.EOF
      SaleID$ = rsWork("Reference")
      rsSales.MoveFirst
      rsSales.Find "[AR SALE Ext Document #]='" & SaleID$ & "'"
      If rsSales.EOF Then
        'Should not happen
      Else
        'Update Sales Record
        If rsWork("Applied Amount") > 0 Then
          'rsSales.Edit
            rsSales("AR SALE Amount Paid") = rsSales("AR SALE Total") - rsWork("Balance")
            rsSales("AR SALE Balance Due") = rsWork("Balance")
          rsSales.Update
        End If
        If rsWork("Applied Amount") > 0 Then 'rsWork("Discount") > 0 Or rsWork("Write Off") > 0 Or rsWork("Applied Amount") > 0 Then
            'Write a cross reference record
            SQLstatement = "INSERT INTO [AR Payment Invoice Cross Reference]"
            SQLstatement = SQLstatement & " ([AR CROSS Payment ID],[AR CROSS Payed ID],[AR CROSS Discount Taken],"
            SQLstatement = SQLstatement & "[AR CROSS Write Off Amount],[AR CROSS Applied Amount],[AR CROSS Cleared])"
            SQLstatement = SQLstatement & " VALUES (" & txtfields(2).Text & "," & rsSales("AR SALE Document #") & ","
            SQLstatement = SQLstatement & rsWork("Discount") & "," & rsWork("Write Off") & "," & rsWork("Applied Amount") & ",False)"
            'Debug.Print SQLstatement
            
            db.Execute SQLstatement
          'rsCross.AddNew
          '  rsCross("AR CROSS Payment ID") = txtFields(2)
          '  rsCross("AR CROSS Payed ID") = rsSales("AR SALE Document #")
          '  rsCross("AR CROSS Discount Taken") = rsWork("Discount")
          '  rsCross("AR CROSS Write Off Amount") = rsWork("Write Off")
          '  rsCross("AR CROSS Applied Amount") = rsWork("Applied Amount")
          '  rsCross("AR CROSS Cleared") = False
          'rsCross.Update
        Else
        End If
      End If
      rsWork.MoveNext
    Loop
  'End If

rsWork.Close
rsSales.Close
Set rsSales = Nothing
Set rsWork = Nothing

Exit Sub
PostData_Error:
  Call LogError("Cash Receipts", "PostData", Now, Err, Error, True)
  Resume Next

End Sub

Private Sub PostDataNSF()
  
'  On Error GoTo PostDataNSF_Error
  
  Dim rsWork As ADODB.Recordset
  'Dim rsCross As ADODB.Recordset
  Dim rsSales As ADODB.Recordset
  Dim SaleID$

  If txtfields(2) = "" Then Exit Sub

  Set rsWork = New Recordset
  rsWork.Open "SELECT * FROM [Cash Receipts Work]", db, adOpenStatic, adLockOptimistic
  
  'Set rsCross = New Recordset
  'rsCross.Open "SELECT * FROM [AR Payment Invoice Cross Reference]", db, adOpenStatic, adLockOptimistic
  
  Set rsSales = New Recordset
  rsSales.Open "SELECT [AR SALE Amount Paid],[AR SALE Balance Due],[AR SALE Document #],[AR SALE Ext Document #] FROM [AR Sales]", db, adOpenStatic, adLockOptimistic

  'rsSales.Index = "Ext Document #"

  'Scratch detail
  db.Execute "DELETE * FROM [AR Payment Invoice Cross Reference] where [AR CROSS Payment ID] = " & txtfields(2)

  'Save detail information
  'On Error Resume Next
  rsWork.MoveFirst
  'If Err = 0 Then
    'On Error GoTo PostDataNSF_Error
    Do While Not rsWork.EOF
      SaleID$ = rsWork("Reference")
      rsSales.MoveFirst
      rsSales.Find "[AR SALE Ext Document #]='" & SaleID$ & "'"
      If rsSales.EOF Then
        'Should not happen
      Else
        'Update Sales Record
        'rsSales.Edit
          rsSales("AR SALE Amount Paid") = rsSales("AR SALE Amount Paid") - rsWork("Applied Amount") - rsWork("Discount") - rsWork("Write Off")
          rsSales("AR SALE Balance Due") = rsSales("AR SALE Balance Due") + rsWork("Applied Amount") + rsWork("Discount") + rsWork("Write Off")
        rsSales.Update
        
        If rsWork("Discount") > 0 Or rsWork("Write Off") > 0 Or rsWork("Applied Amount") > 0 Then
      
            'Write a cross reference record
            SQLstatement = "INSERT INTO [AR Payment Invoice Cross Reference]"
            SQLstatement = SQLstatement & " ([AR CROSS Payment ID],[AR CROSS Payed ID],[AR CROSS Discount Taken],"
            SQLstatement = SQLstatement & "[AR CROSS Write Off Amount],[AR CROSS Applied Amount],[AR CROSS Cleared])"
            SQLstatement = SQLstatement & " VALUES (" & txtfields(2).Text & "," & rsSales("AR SALE Document #") & ","
            SQLstatement = SQLstatement & rsWork("Discount") & "," & rsWork("Write Off") & "," & rsWork("Applied Amount") & ",False)"
            
            db.Execute SQLstatement
   '       rsCross.AddNew
   '         rsCross("AR CROSS Payment ID") = txtFields(2)
   '         rsCross("AR CROSS Payed ID") = rsSales("AR SALE Document #")
   '         rsCross("AR CROSS Discount Taken") = rsWork("Discount")
   '         rsCross("AR CROSS Write Off Amount") = rsWork("Write Off")
   '         rsCross("AR CROSS Applied Amount") = rsWork("Applied Amount")
   '         rsCross("AR CROSS Cleared") = False
   '       rsCross.Update
        Else
        End If
      End If
      rsWork.MoveNext
    Loop
  'End If

rsWork.Close
rsSales.Close
Set rsSales = Nothing
Set rsWork = Nothing

  Exit Sub
PostDataNSF_Error:
  Call LogError("Cash Receipts", "PostDataNSF", Now, Err, Error, True)
  Resume Next

End Sub

Private Function PostNSF() As Integer

'  On Error GoTo PostNSF_Error
  
  Dim CurrentBalance@
  Dim msg$
  Dim title$

  ' don't mark NSF if not posted
  'If Me![AR PAY Posted YN] = False Then Exit Function
  If chkFields(9).Value = 0 Then
     MsgBox "This is not a posted transaction", vbInformation, "Information"
     Exit Function
  End If
  
  Dim rsCompany As ADODB.Recordset
  Set rsCompany = New Recordset
  rsCompany.Open "SELECT * FROM [SYS Company]", db, adOpenStatic, adLockOptimistic
  rsCompany.MoveFirst
  
  'Dim rsGLWorkDetailTemp As ADODB.Recordset
  'Set rsGLWorkDetailTemp = New Recordset
  'rsGLWorkDetailTemp.Open "SELECT * FROM [GL Work Detail]", db, adOpenStatic, adLockOptimistic

  'Post by 1 - system date or 2 - Transaction date?
  Dim PostDate%
  PostDate% = rsCompany("SYS COM GL Post By Date")

  'Set Invoice Type
  Dim TranDate As Variant

  'Set Post Date
  If PostDate% = 1 Then
    TranDate = DateValue(Format(Now, "Short Date"))
  Else
    'TranDate = DateValue(Me![AR PAY Transaction Date])
    TranDate = DateValue(txtfields(3))
  End If
  
  'Verify period can be posted to
  'Send TranDate
  'Return PeriodToPost and PeriodClosed
  Dim PeriodToPost%
  Dim PeriodClosed%
  
  Call VerifyPeriod(TranDate, PeriodToPost%, PeriodClosed%)
  
  If PeriodClosed% = True Then
    MsgBox "Unable to post transaction to a Closed Period.", , "Post Payment Error"
    PostNSF% = False
    Exit Function
  End If

  ' save the datafirst
  Call PostDataNSF

  ' clear any GL Work records
  'MsgBox db.ConnectionString
  db.Execute "DELETE DISTINCTROW * FROM [GL Work Detail]"

  ' write GL Transaction Header
  Dim refr$
  Dim desc$
  Dim NewNumber&
  Dim rsCustomer As ADODB.Recordset
  Set rsCustomer = New Recordset
  rsCustomer.Open "SELECT * FROM [AR Customer] WHERE [AR CUST Customer ID]='" & txtfields(0).Text & "'", db, adOpenStatic, adLockOptimistic, adCmdText
  'rsCustomer.Index = "PrimaryKey"
  'rsCustomer.MoveFirst
  'rsCustomer.Find "[AR CUST Customer ID]='" & txtFields(0).Text & "'"

  db.Execute "DELETE * FROM [GL Transaction] WHERE [GL TRANS Document #]='" & "CASH REC " & txtfields(5) & "-" & txtfields(0) & "'"
  'rsGLTrans.AddNew
'again:
    'rsGLTrans("GL TRANS Document #") = "NSF " & txtFields(5)
    'rsGLTrans("GL TRANS Type") = "NSF"
    Dim SQLstatement As String
    ' gl post date
      SQLstatement = "INSERT INTO [GL Transaction]"
      SQLstatement = SQLstatement & " ([GL TRANS Document #],[GL TRANS Type],[GL TRANS Date],"
      SQLstatement = SQLstatement & " [GL TRANS Reference],[GL TRANS Amount],[GL TRANS Posted YN],"
      SQLstatement = SQLstatement & " [GL TRANS Description],[GL TRANS Source],[GL TRANS System Generated])"
    
    Dim TempStr As String
    
    If PostDate% = 1 Then
      TempStr = Format(Now, "Short Date")
    Else
      TempStr = txtfields(3).Text 'Me![AR PAY Transaction Date]
    End If
    
      SQLstatement = SQLstatement & " VALUES ('NSF " & txtfields(5) & "','NSF',#" & TempStr & "#,"
      
    If rsCustomer.EOF Then
      refr$ = "Unknown"
    Else
      refr$ = rsCustomer("AR CUST Name")
    End If
      
      SQLstatement = SQLstatement & "'" & refr$ & "'," & CCur(txtfields(4).Text) & ",1,"
      SQLstatement = SQLstatement & "'NSF " & txtfields(5).Text & "','NSF " & txtfields(5).Text & "',True)"
      'Debug.Print SQLstatement
      
      db.Execute SQLstatement
      'GoTo again
      
    'rsGLTrans("GL TRANS Reference") = refr$
    'rsGLTrans("GL TRANS Amount") = txtFields(4).Text 'Me![AR PAY Amount]
    'rsGLTrans("GL TRANS Posted YN") = 1
    'desc$ = "NSF " & txtFields(5).Text 'Me![AR PAY Check No]
    'rsGLTrans("GL TRANS Description") = desc$
    'rsGLTrans("GL TRANS Source") = "NSF " & txtFields(5).Text
    'rsGLTrans("GL TRANS System Generated") = True
  'rsGLTrans.Update
  'write GL Transaction Detail
  
  Dim rsGLTrans As ADODB.Recordset
  Set rsGLTrans = New Recordset
  rsGLTrans.Open "SELECT [GL TRANS Number] FROM [GL Transaction]WHERE [GL TRANS Document #]='" & "NSF " & txtfields(5).Text & "'", db, adOpenStatic, adLockOptimistic, adCmdText
      NewNumber& = rsGLTrans("GL TRANS Number")
  rsGLTrans.Close
  Set rsGLTrans = Nothing
  
  'Loop through line items
  Dim rsCross As ADODB.Recordset
  Set rsCross = New Recordset
  rsCross.Open "SELECT * FROM [AR Payment Invoice Cross Reference] WHERE [AR CROSS Payment ID] = " & txtfields(2), db, adOpenStatic, adLockOptimistic, adCmdText

  'On Error Resume Next
  If rsCross.RecordCount > 0 Then
  rsCross.MoveFirst
    Do While Not rsCross.EOF
      ' only process records with payments, discounts or writeoffs
      If rsCross("AR CROSS Applied Amount") > 0 Or rsCross("AR CROSS Discount Taken") > 0 Or rsCross("AR CROSS Write Off Amount") > 0 Then
        
        ' process payments
        If rsCross("AR CROSS Applied Amount") > 0 Then
          
'          ' update GL for payment
          '-----------------------------------------------------------------------
          ' Payment GL Affected Accounts
          '
          '                  Debit   Credit   Source
          '                  -----   ------   ------
          ' AR                 X              Pref - Sales
          ' CASH                       X      Bank - Cash Acct
          '-----------------------------------------------------------------------

          ' Debits
          ' AR
      SQLstatement = "INSERT INTO [GL Work Detail]"
      SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
      SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & rsCompany("SYS COM Sales AR Acct") & "" & "'," & rsCross("AR CROSS Applied Amount") & ",0)"
      db.Execute SQLstatement

      '   rsGLWorkDetailTemp.AddNew
      '      rsGLWorkDetailTemp("GW TRANSD Number") = NewNumber&
      '      rsGLWorkDetailTemp("GW TRANSD Account") = rsCompany("SYS COM Sales AR Acct") & ""
      '      rsGLWorkDetailTemp("GW TRANSD Debit Amount") = rsCross("AR CROSS Applied Amount")
      '      rsGLWorkDetailTemp("GW TRANSD Credit Amount") = 0
      '      rsGLWorkDetailTemp("GW TRANSD Project") = ""
      '    rsGLWorkDetailTemp.Update
          ' update GL for payment

          ' Credits
          ' NSF
      SQLstatement = "INSERT INTO [GL Work Detail]"
      SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
      SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & txtfields(1).Text & "" & "',0," & rsCross("AR CROSS Applied Amount") & ")"
      db.Execute SQLstatement
          
      '    rsGLWorkDetailTemp.AddNew
      '      rsGLWorkDetailTemp("GW TRANSD Number") = NewNumber&
      '      rsGLWorkDetailTemp("GW TRANSD Account") = txtFields(1).Text ' Me![AR PAY Bank Account]
      '      rsGLWorkDetailTemp("GW TRANSD Debit Amount") = 0
      '      rsGLWorkDetailTemp("GW TRANSD Credit Amount") = rsCross("AR CROSS Applied Amount")
      '      rsGLWorkDetailTemp("GW TRANSD Project") = ""
      '    rsGLWorkDetailTemp.Update
          
        End If ' end process payments
        
        ' process discount amounts
        If rsCross("AR CROSS Discount Taken") > 0 Then

'          ' update GL for discount
          '-----------------------------------------------------------------------
          ' Discount GL Affected Accounts
          '
          '                  Debit   Credit   Source
          '                  -----   ------   ------
          ' AR                 X              Pref - Sales
          ' Discount                   X       Pref - Sales
          '-----------------------------------------------------------------------

          ' Debits
          ' AR
      SQLstatement = "INSERT INTO [GL Work Detail]"
      SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
      SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & rsCompany("SYS COM Sales AR Acct") & "" & "'," & rsCross("AR CROSS Discount Taken") & ",0)"
      db.Execute SQLstatement
          
      '    rsGLWorkDetailTemp.AddNew
      '      rsGLWorkDetailTemp("GW TRANSD Number") = NewNumber&
      '      rsGLWorkDetailTemp("GW TRANSD Account") = rsCompany("SYS COM Sales AR Acct")
      '      rsGLWorkDetailTemp("GW TRANSD Debit Amount") = rsCross("AR CROSS Discount Taken")
      '      rsGLWorkDetailTemp("GW TRANSD Credit Amount") = 0
      '      rsGLWorkDetailTemp("GW TRANSD Project") = ""
      '    rsGLWorkDetailTemp.Update
          ' update GL for discount

          ' Credits
          ' Discount
      SQLstatement = "INSERT INTO [GL Work Detail]"
      SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
      SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & rsCompany("SYS COM Sales Discount Acct") & "" & "',0," & rsCross("AR CROSS Discount Taken") & ")"
      db.Execute SQLstatement
          
      '    rsGLWorkDetailTemp.AddNew
      '      rsGLWorkDetailTemp("GW TRANSD Number") = NewNumber&
      '      rsGLWorkDetailTemp("GW TRANSD Account") = rsCompany("SYS COM Sales Discount Acct")
      '      rsGLWorkDetailTemp("GW TRANSD Debit Amount") = 0
      '      rsGLWorkDetailTemp("GW TRANSD Credit Amount") = rsCross("AR CROSS Discount Taken")
      '      rsGLWorkDetailTemp("GW TRANSD Project") = ""
      '    rsGLWorkDetailTemp.Update

        End If ' end process discount amounts
        
        ' process write off amounts
        If rsCross("AR CROSS Write Off Amount") > 0 Then

'          ' update GL for discount
          '-----------------------------------------------------------------------
          ' Write Off GL Affected Accounts
          '
          '                  Debit   Credit   Source
          '                  -----   ------   ------
          ' AR                         X      Pref - Sales
          ' WriteOff           X              Pref - Sales
          '-----------------------------------------------------------------------

          ' Debits
          ' AR
      SQLstatement = "INSERT INTO [GL Work Detail]"
      SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
      SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & rsCompany("SYS COM Sales AR Acct") & "" & "'," & rsCross("AR CROSS Write Off Amount") & ",0)"
      db.Execute SQLstatement
          
      '    rsGLWorkDetailTemp.AddNew
      '      rsGLWorkDetailTemp("GW TRANSD Number") = NewNumber&
      '      rsGLWorkDetailTemp("GW TRANSD Account") = rsCompany("SYS COM Sales AR Acct")
      '      rsGLWorkDetailTemp("GW TRANSD Debit Amount") = rsCross("AR CROSS Write Off Amount")
      '      rsGLWorkDetailTemp("GW TRANSD Credit Amount") = 0
      '      rsGLWorkDetailTemp("GW TRANSD Project") = ""
      '    rsGLWorkDetailTemp.Update
          ' update GL for discount

          ' Credits
          ' Write Off
      SQLstatement = "INSERT INTO [GL Work Detail]"
      SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
      SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & rsCompany("SYS COM Sales Write Off Acct") & "" & "',0," & rsCross("AR CROSS Write Off Amount") & ")"
      db.Execute SQLstatement
          
     '     rsGLWorkDetailTemp.AddNew
     '       rsGLWorkDetailTemp("GW TRANSD Number") = NewNumber&
     '       rsGLWorkDetailTemp("GW TRANSD Account") = rsCompany("SYS COM Sales Write Off Acct")
     '       rsGLWorkDetailTemp("GW TRANSD Debit Amount") = 0
     '       rsGLWorkDetailTemp("GW TRANSD Credit Amount") = rsCross("AR CROSS Write Off Amount")
     '       rsGLWorkDetailTemp("GW TRANSD Project") = ""
     '     rsGLWorkDetailTemp.Update

        End If ' end process write offs
      End If
      rsCross.MoveNext
    Loop
  End If

  ' handle Unapplied Payments or Payments On Account
'  If Forms![Cash Receipts].[Cash Receipts Detail].Form![Unapplied Amount] > 0 Then
  If txtfields(12) > 0 Then
    ' Debits
    ' AR
      SQLstatement = "INSERT INTO [GL Work Detail]"
      SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
      SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & rsCompany("SYS COM Sales AR Acct") & "" & "'," & CCur(txtfields(12).Text) & ",0)"
      db.Execute SQLstatement
          
    'rsGLWorkDetailTemp.AddNew
    '  rsGLWorkDetailTemp("GW TRANSD Number") = NewNumber&
    '  rsGLWorkDetailTemp("GW TRANSD Account") = rsCompany("SYS COM Sales AR Acct")
    '  rsGLWorkDetailTemp("GW TRANSD Debit Amount") = txtFields(12) 'Forms![Cash Receipts].[Cash Receipts Detail].Form![Unapplied Amount]
    '  rsGLWorkDetailTemp("GW TRANSD Credit Amount") = 0
    '  rsGLWorkDetailTemp("GW TRANSD Project") = ""
    'rsGLWorkDetailTemp.Update
    ' update GL for payment

    ' Credits
    ' Cash Receipt
      SQLstatement = "INSERT INTO [GL Work Detail]"
      SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
      SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & txtfields(1).Text & "',0," & CCur(txtfields(12).Text) & ")"
      db.Execute SQLstatement
    
    'rsGLWorkDetailTemp.AddNew
    '  rsGLWorkDetailTemp("GW TRANSD Number") = NewNumber&
    '  rsGLWorkDetailTemp("GW TRANSD Account") = txtFields(1).text
    '  rsGLWorkDetailTemp("GW TRANSD Debit Amount") = 0
    '  rsGLWorkDetailTemp("GW TRANSD Credit Amount") = txtFields(12).text 'Forms![Cash Receipts].[Cash Receipts Detail].Form![Unapplied Amount]
    '  rsGLWorkDetailTemp("GW TRANSD Project") = ""
    'rsGLWorkDetailTemp.Update

  End If
  ' end of handle Unapplied Payments or Payments On Account

  ' update customer stats
  'rsCustomer.Index = "PrimaryKey"
  'rsCustomer.MoveFirst
  'rsCustomer.Find "[AR CUST Customer ID]='" & txtFields(0).Text & "'"
  If rsCustomer.EOF Then
  Else
    'rsCustomer.Edit
      rsCustomer("AR CUST Payments YTD") = rsCustomer("AR CUST Payments YTD") - txtfields(4)
      rsCustomer("AR CUST Payments Lifetime") = rsCustomer("AR CUST Payments Lifetime") - txtfields(4)
      ' Update current Balance - if not Paid in full
      CurrentBalance@ = IIf(IsNull(rsCustomer("AR CUST Financial Period 1")), 0, rsCustomer("AR CUST Financial Period 1"))
      CurrentBalance@ = CurrentBalance@ + txtfields(4)
      rsCustomer("AR CUST Financial Period 1") = CurrentBalance@
    rsCustomer.Update
  End If

  'Post GL Entry
  Dim Success%
  Success% = PostGLWorkDetail(TranDate, NewNumber&, db)

  If Success% = False Then
    PostNSF = False
  Else
    PostNSF = True
  End If

PostNSF_Exit:

  Exit Function

PostNSF_Error:
  Call LogError("Cash Receipts", "PostNSF", Now, Err, Error, True)
  PostNSF = False
  Exit Function
  Resume Next

End Function

Private Function AR_PostPayments() As Integer
'On Error GoTo AR_PostPayments_Error
  
  Dim CurrentBalance@
  Dim msg$
  Dim title$

  ' don't post if already posted
  If chkFields(9).Value = 1 Then Exit Function
  
  Dim rsCompany As ADODB.Recordset
  Set rsCompany = New Recordset
  rsCompany.Open "SELECT  [SYS COM Sales AR Acct],[SYS COM Sales Discount Acct]," & _
  "[SYS COM Sales AR Acct],[SYS COM Sales Write Off Acct],[SYS COM GL Post By Date] " & _
  "FROM [SYS Company]", db, adOpenStatic, adLockOptimistic
  rsCompany.MoveFirst
  
  'Dim rsGLWorkDetail As ADODB.Recordset
  'Set rsGLWorkDetail = New Recordset
  'rsGLWorkDetail.Open "SELECT * FROM [GL Work Detail]", db, adOpenStatic, adLockOptimistic

  'Post by 1 - system date or 2 - Transaction date?
  Dim PostDate%
  PostDate% = rsCompany("SYS COM GL Post By Date")

  'Set Invoice Type
  Dim TranDate As Variant

  'Set Post Date
  If PostDate% = 1 Then
    TranDate = DateValue(Format(Now, "Short Date"))
  Else
    TranDate = DateValue(txtfields(3))
  End If
  
  'Verify period can be posted to
  'Send TranDate
  'Return PeriodToPost and PeriodClosed
  Dim PeriodToPost%
  Dim PeriodClosed%
  
  Call VerifyPeriod(TranDate, PeriodToPost%, PeriodClosed%)
  
  If PeriodClosed% = True Then
    MsgBox "Unable to post transaction to a Closed Period.", , "Post Payment Error"
    AR_PostPayments% = False
    Exit Function
  End If

  ' save the datafirst
  Call PostData
    
  ' clear any GL Work records
  ' db.Execute "qryDeleteGLWorkDetail"
  db.Execute "DELETE DISTINCTROW * FROM [GL Work Detail]"
  
  'Dim rsGLTrans As ADODB.Recordset
  'Set rsGLTrans = New Recordset
  'rsGLTrans.Open "SELECT * FROM [GL Transaction]", db, adOpenStatic, adLockOptimistic

  ' write GL Transaction Header
  Dim refr$
  Dim desc$
  Dim NewNumber&
  Dim rsCustomer As ADODB.Recordset
  Set rsCustomer = New Recordset
  rsCustomer.Open "SELECT * FROM [AR Customer]", db, adOpenStatic, adLockOptimistic
  'rsCustomer.Index = "PrimaryKey"
  rsCustomer.MoveFirst
  rsCustomer.Find "[AR CUST Customer ID]='" & txtfields(0) & "'"
  
  'rsGLTrans.AddNew
'    NewNumber& = rsGLTrans("GL TRANS Number")

    Dim SQLstatement As String
    ' gl post date
      SQLstatement = "INSERT INTO [GL Transaction]"
      SQLstatement = SQLstatement & " ([GL TRANS Document #],[GL TRANS Type],[GL TRANS Date],"
      SQLstatement = SQLstatement & " [GL TRANS Reference],[GL TRANS Amount],[GL TRANS Posted YN],"
      SQLstatement = SQLstatement & " [GL TRANS Description],[GL TRANS Source],[GL TRANS System Generated])"
    
    Dim TempStr As String
    'xxx 3/26/97 v7.3 Changed Account No to Customer ID
    'rsGLTrans("GL TRANS Document #") = "CASH REC " & txtFields(5) & "-" & txtFields(0).Text
  'Debug.Print "DELETE * FROM [GL Transaction] WHERE [GL TRANS Document #]='" & "CASH REC " & txtFields(5) & "-" & txtFields(0) & "'"
    'rsGLTrans("GL TRANS Type") = "Cash Receipt"
    
    ' gl post date
    If PostDate% = 1 Then
      TempStr = Format(Now, "Short Date")
    Else
      TempStr = txtfields(3).Text
    End If
      
      SQLstatement = SQLstatement & " VALUES ('CASH REC " & txtfields(5) & "-" & txtfields(0).Text & "','Cash Receipt',#" & TempStr & "#,"
    
    If rsCustomer.EOF Then
      refr$ = "Unknown"
    Else
      refr$ = rsCustomer("AR CUST Name")
    End If
      
      SQLstatement = SQLstatement & "'" & refr$ & "'," & CCur(txtfields(4).Text) & ",1,"
      SQLstatement = SQLstatement & "'CASH REC " & txtfields(5).Text & "','CASH REC " & txtfields(5).Text & "',True)"
      'Debug.Print SQLstatement
      
  db.Execute "DELETE * FROM [GL Transaction] WHERE [GL TRANS Document #]='" & "CASH REC " & txtfields(5) & "-" & txtfields(0) & "'"
  db.Execute SQLstatement

 '   rsGLTrans("GL TRANS Reference") = refr$
 '   rsGLTrans("GL TRANS Amount") = txtFields(4)
 '   rsGLTrans("GL TRANS Posted YN") = 1
 '   desc$ = "CASH REC " & txtFields(5)
 '   rsGLTrans("GL TRANS Description") = desc$
 '   rsGLTrans("GL TRANS Source") = "CASH REC " & txtFields(5)
 '   rsGLTrans("GL TRANS System Generated") = True
 ' rsGLTrans.Update
  Dim rsGLTrans As ADODB.Recordset
  Set rsGLTrans = New Recordset
  rsGLTrans.Open "SELECT [GL TRANS Number] FROM [GL Transaction]WHERE [GL TRANS Document #]='" & "CASH REC " & txtfields(5) & "-" & txtfields(0).Text & "'", db, adOpenStatic, adLockOptimistic, adCmdText
      NewNumber& = rsGLTrans("GL TRANS Number")
  rsGLTrans.Close
  Set rsGLTrans = Nothing
  ' write GL Transaction Detail

  'Loop through line items
  Dim rsCross As ADODB.Recordset
  Set rsCross = New Recordset
  rsCross.Open "SELECT * FROM [AR Payment Invoice Cross Reference] WHERE [AR CROSS Payment ID] = " & txtfields(2), db, adOpenStatic, adLockOptimistic, adCmdText
  'rsCross.Filter = "[AR CROSS Payment ID] = '" & txtFields(2) & "'"
  
  'On Error Resume Next
  'xxx 4/15/97 7.3 rsCross.movelast used below
  'rsCross.MoveLast

  'Debug.Print "d"
  If rsCross.RecordCount > 0 Then
    rsCross.MoveFirst
    Do While Not rsCross.EOF
      ' only process records with payments, discounts or writeoffs
      If rsCross("AR CROSS Applied Amount") > 0 Or rsCross("AR CROSS Discount Taken") > 0 Or rsCross("AR CROSS Write Off Amount") > 0 Then
        
        ' process payments
        If rsCross("AR CROSS Applied Amount") > 0 Then
          
'          ' update GL for payment
          '-----------------------------------------------------------------------
          ' Payment GL Affected Accounts
          '
          '                  Debit   Credit   Source
          '                  -----   ------   ------
          ' CASH               X              Bank - Cash Acct
          ' AR                         X      Pref - Sales
          '-----------------------------------------------------------------------

          ' Debits
          ' Cash Receipt
      SQLstatement = "INSERT INTO [GL Work Detail]"
      SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
      SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & txtfields(1).Text & "'," & rsCross("AR CROSS Applied Amount") & ",0)"
      db.Execute SQLstatement
          
      '    rsGLWorkDetail.AddNew
      '      rsGLWorkDetail("GW TRANSD Number") = NewNumber&
      '      rsGLWorkDetail("GW TRANSD Account") = txtFields(1).Text
      '      rsGLWorkDetail("GW TRANSD Debit Amount") = rsCross("AR CROSS Applied Amount")
      '      rsGLWorkDetail("GW TRANSD Credit Amount") = 0
      '      rsGLWorkDetail("GW TRANSD Project") = ""
      '    rsGLWorkDetail.Update

          ' Credits
          ' AR
      
      SQLstatement = "INSERT INTO [GL Work Detail]"
      SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
      SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & rsCompany("SYS COM Sales AR Acct") & "" & "',0," & rsCross("AR CROSS Applied Amount") & ")"
      db.Execute SQLstatement
          
      '    rsGLWorkDetail.AddNew
      '      rsGLWorkDetail("GW TRANSD Number") = NewNumber&
      '      rsGLWorkDetail("GW TRANSD Account") = rsCompany("SYS COM Sales AR Acct")
      '      rsGLWorkDetail("GW TRANSD Debit Amount") = 0
      '      rsGLWorkDetail("GW TRANSD Credit Amount") = rsCross("AR CROSS Applied Amount")
      '      rsGLWorkDetail("GW TRANSD Project") = ""
      '    rsGLWorkDetail.Update
          ' update GL for payment

        End If ' end process payments
        
        ' process discount amounts
        If rsCross("AR CROSS Discount Taken") > 0 Then

 '         ' update GL for discount
          '-----------------------------------------------------------------------
          ' Discount GL Affected Accounts
          '
          '                  Debit   Credit   Source
          '                  -----   ------   ------
          ' Discount           X              Pref - Sales
          ' AR                         X      Pref - Sales
          '-----------------------------------------------------------------------

          ' Debits
          ' Discount
          
      SQLstatement = "INSERT INTO [GL Work Detail]"
      SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
      SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & rsCompany("SYS COM Sales Discount Acct") & "" & "'," & rsCross("AR CROSS Discount Taken") & ",0)"
      db.Execute SQLstatement
          
      '    rsGLWorkDetail.AddNew
      '      rsGLWorkDetail("GW TRANSD Number") = NewNumber&
      '      rsGLWorkDetail("GW TRANSD Account") = rsCompany("SYS COM Sales Discount Acct")
      '      rsGLWorkDetail("GW TRANSD Debit Amount") = rsCross("AR CROSS Discount Taken")
      '      rsGLWorkDetail("GW TRANSD Credit Amount") = 0
      '      rsGLWorkDetail("GW TRANSD Project") = ""
      '    rsGLWorkDetail.Update

          ' Credits
          ' AR
          
      SQLstatement = "INSERT INTO [GL Work Detail]"
      SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
      SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & rsCompany("SYS COM Sales AR Acct") & "" & "',0," & rsCross("AR CROSS Discount Taken") & ")"
      db.Execute SQLstatement
          
      '    rsGLWorkDetail.AddNew
      '      rsGLWorkDetail("GW TRANSD Number") = NewNumber&
      '      rsGLWorkDetail("GW TRANSD Account") = rsCompany("SYS COM Sales AR Acct")
      '      rsGLWorkDetail("GW TRANSD Debit Amount") = 0
      '      rsGLWorkDetail("GW TRANSD Credit Amount") = rsCross("AR CROSS Discount Taken")
      '      rsGLWorkDetail("GW TRANSD Project") = ""
      '    rsGLWorkDetail.Update
          ' update GL for discount

        End If ' end process discount amounts
        
        ' process write off amounts
        If rsCross("AR CROSS Write Off Amount") > 0 Then

'          ' update GL for discount
          '-----------------------------------------------------------------------
          ' Write Off GL Affected Accounts
          '
          '                  Debit   Credit   Source
          '                  -----   ------   ------
          ' WriteOff           X              Pref - Sales
          ' AR                         X      Pref - Sales
          '-----------------------------------------------------------------------

          ' Debits
          ' Write Off
          
      SQLstatement = "INSERT INTO [GL Work Detail]"
      SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
      SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & rsCompany("SYS COM Sales Write Off Acct") & "" & "'," & rsCross("AR CROSS Write Off Amount") & ",0)"
      db.Execute SQLstatement
          
      '    rsGLWorkDetail.AddNew
      '      rsGLWorkDetail("GW TRANSD Number") = NewNumber&
      '      rsGLWorkDetail("GW TRANSD Account") = rsCompany("SYS COM Sales Write Off Acct")
      '      rsGLWorkDetail("GW TRANSD Debit Amount") = rsCross("AR CROSS Write Off Amount")
      '      rsGLWorkDetail("GW TRANSD Credit Amount") = 0
      '      rsGLWorkDetail("GW TRANSD Project") = ""
      '    rsGLWorkDetail.Update

          ' Credits
          ' AR
          
      SQLstatement = "INSERT INTO [GL Work Detail]"
      SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
      SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & rsCompany("SYS COM Sales AR Acct") & "" & "',0," & rsCross("AR CROSS Write Off Amount") & ")"
      db.Execute SQLstatement
          
      '    rsGLWorkDetail.AddNew
      '      rsGLWorkDetail("GW TRANSD Number") = NewNumber&
      '      rsGLWorkDetail("GW TRANSD Account") = rsCompany("SYS COM Sales AR Acct")
      '      rsGLWorkDetail("GW TRANSD Debit Amount") = 0
      '      rsGLWorkDetail("GW TRANSD Credit Amount") = rsCross("AR CROSS Write Off Amount")
      '      rsGLWorkDetail("GW TRANSD Project") = ""
      '    rsGLWorkDetail.Update
          ' update GL for discount

        End If ' end process write offs
      End If
      rsCross.MoveNext
    Loop
  End If
  
  'xxx 4/15/97 Added up to ElseIf (ElseIf was just If)
  ' handle Unapplied Payments or Payments On Account
  If rsCross.RecordCount = 0 Then
    ' Debits
    ' Cash Receipt
      SQLstatement = "INSERT INTO [GL Work Detail]"
      SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
      SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & txtfields(1).Text & "'," & CCur(txtfields(4).Text) & ",0)"
      db.Execute SQLstatement
          
   'rsGLWorkDetail.AddNew
   '   rsGLWorkDetail("GW TRANSD Number") = NewNumber&
   '   rsGLWorkDetail("GW TRANSD Account") = txtFields(1).Text
   '   rsGLWorkDetail("GW TRANSD Debit Amount") = txtFields(4).Text
   '   rsGLWorkDetail("GW TRANSD Credit Amount") = 0
   '   rsGLWorkDetail("GW TRANSD Project") = ""
   ' rsGLWorkDetail.Update

    ' Credits
    ' AR
      SQLstatement = "INSERT INTO [GL Work Detail]"
      SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
      SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & rsCompany("SYS COM Sales AR Acct") & "',0," & CCur(txtfields(4).Text) & ")"
      db.Execute SQLstatement
    
    'rsGLWorkDetail.AddNew
    '  rsGLWorkDetail("GW TRANSD Number") = NewNumber&
    '  rsGLWorkDetail("GW TRANSD Account") = rsCompany("SYS COM Sales AR Acct")
    '  rsGLWorkDetail("GW TRANSD Debit Amount") = 0
    '  rsGLWorkDetail("GW TRANSD Credit Amount") = txtFields(4)
    '  rsGLWorkDetail("GW TRANSD Project") = ""
    'rsGLWorkDetail.Update
    ' update GL for payment
  'ElseIf Forms![Cash Receipts].[Cash Receipts Detail].Form![Unapplied Amount] > 0 Then
  '=[Forms]![Cash Receipts].[AR PAY Amount]-Sum([Applied Amount])
  ElseIf CCur(txtfields(12).Text) > 0 Then

    ' Debits
    ' Cash Receipt
      SQLstatement = "INSERT INTO [GL Work Detail]"
      SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
      SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & txtfields(1).Text & "'," & CCur(txtfields(12).Text) & ",0)"
      db.Execute SQLstatement
    
    'rsGLWorkDetail.AddNew
    '  rsGLWorkDetail("GW TRANSD Number") = NewNumber&
    '  rsGLWorkDetail("GW TRANSD Account") = txtFields(1).Text
    '  rsGLWorkDetail("GW TRANSD Debit Amount") = txtFields(12).Text
    '  rsGLWorkDetail("GW TRANSD Credit Amount") = 0
    '  rsGLWorkDetail("GW TRANSD Project") = ""
    'rsGLWorkDetail.Update

    ' Credits
    ' AR
      SQLstatement = "INSERT INTO [GL Work Detail]"
      SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
      SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & rsCompany("SYS COM Sales AR Acct") & "',0," & CCur(txtfields(12).Text) & ")"
      db.Execute SQLstatement
    
    'rsGLWorkDetail.AddNew
    '  rsGLWorkDetail("GW TRANSD Number") = NewNumber&
    '  rsGLWorkDetail("GW TRANSD Account") = rsCompany("SYS COM Sales AR Acct")
    '  rsGLWorkDetail("GW TRANSD Debit Amount") = 0
    '  rsGLWorkDetail("GW TRANSD Credit Amount") = txtFields(12).Text
    '  rsGLWorkDetail("GW TRANSD Project") = ""
    'rsGLWorkDetail.Update
    ' update GL for payment
  
  End If
  ' end of handle Unapplied Payments or Payments On Account

  ' update customer stats
  'Set rsCustomer = DtBase2.OpenRecordset("AR Customer")
  
  'rsCustomer.Index = "PrimaryKey"
  'rsCustomer.MoveFirst
  'rsCustomer.Find "[AR CUST Customer ID]='" & txtFields(0).Text & "'"

  'rsCustomer.Edit
    rsCustomer("AR CUST Payments YTD") = rsCustomer("AR CUST Payments YTD") + txtfields(4)
    rsCustomer("AR CUST Payments Lifetime") = rsCustomer("AR CUST Payments Lifetime") + txtfields(4)
    ' Update current Balance - if not Paid in full
    CurrentBalance@ = IIf(IsNull(rsCustomer("AR CUST Financial Period 1")), 0, rsCustomer("AR CUST Financial Period 1"))
    CurrentBalance@ = CurrentBalance@ - txtfields(4)
    rsCustomer("AR CUST Financial Period 1") = CurrentBalance@
  rsCustomer.Update

  ' post GL entry
  Dim Success%
  Success% = PostGLWorkDetail(TranDate, NewNumber&, db)
  If Success% = False Then
    AR_PostPayments% = False
  Else
    AR_PostPayments = True
  End If

AR_PostPayments_Exit:
  
Exit Function

AR_PostPayments_Error:
  Call LogError("Cash Receipts", "AR_PostPayments", Now, Err.Number, Err.Description, True)
  AR_PostPayments% = False
  Exit Function
  'Resume Next

End Function


Private Function CheckAllData() As Boolean
 
  'If the type ID exists then load that record
  'Otherwise create a new record w/ this ID.

  'On Error GoTo AR_PAY_Check_No_AfterUpdate_Error

  'Dim tb As ADODB.Recordset
  Dim HoldID$
  Dim HoldCheck$

   Set ADOprimaryrs = New Recordset
   'Debug.Print "select [AR PAY ID],[AR PAY Type],[AR PAY Check No],[AR PAY Customer No],[AR PAY Transaction Date],[AR PAY Amount],[AR PAY Bank Account],[AR PAY Posted YN],[AR PAY Reconciled],[AR PAY Cleared],[AR PAY Deposited YN],[AR PAY NSF],[AR PAY UnApplied Amount],[AR PAY Status],[AR PAY Notes] from [AR Payment Header] WHERE [AR PAY Customer No]='" & txtfields(0) & "'AND [AR PAY Check No]='" & txtfields(5) & "'"
   ADOprimaryrs.Open "select [AR PAY ID],[AR PAY Type],[AR PAY Check No],[AR PAY Customer No]," & _
   "[AR PAY Transaction Date],[AR PAY Amount],[AR PAY Bank Account],[AR PAY Posted YN]," & _
   "[AR PAY Reconciled],[AR PAY Cleared],[AR PAY Deposited YN],[AR PAY NSF],[AR PAY UnApplied Amount]," & _
   "[AR PAY Status],[AR PAY Notes] from [AR Payment Header] " & _
   "WHERE [AR PAY Customer No]='" & txtfields(0) & "'AND [AR PAY Check No]='" & _
   txtfields(5) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
  
  HoldID$ = txtfields(0)
  HoldCheck$ = txtfields(5)
  
  With ADOprimaryrs
    If ADOprimaryrs.RecordCount = 0 Then
        'Dim CreateReceipt As Integer
        grdDatagrid.SetFocus
        'If grdDataGrid.Row = -1 Then
        '    MsgBox "You have to select a transaction to continue", vbInformation, "Error"
        '    txtFields(5) = ""
        '    CheckAllData = False
        '    Exit Function
        'Else
        '    CreateReceipt = MsgBox("Attempting to create a new Cash Receipt called " & txtFields(5) & " on" & vbCr & _
        '    "Reference number " & grdDataGrid.Columns(0) & " and " & grdDataGrid.Columns(7) & " in balance" & vbCr & _
        '    "Would you like to create a new Cash Receipt?" _
        '                    , vbYesNo, "Create Receipt")
        '    If CreateReceipt = vbNo Then
        '        txtFields(3) = ""
        '        CheckAllData = False
        '        Exit Function
        '    End If
        'End If
        
        For Each Ctrl In Me.Controls
           If TypeOf Ctrl Is TextBox Or TypeOf Ctrl Is CheckBox Then
              Set Ctrl.DataSource = ADOprimaryrs
                
              If TypeOf Ctrl Is TextBox And Ctrl.DataField <> "" Then
                 If ADOprimaryrs("" & Ctrl.DataField & "").Type = 202 Then Ctrl.MaxLength = ADOprimaryrs("" & Ctrl.DataField & "").DefinedSize
              End If
           End If
        Next
        CheckAllData = False
        'Create a new record
        cmdbankAccount.Enabled = True
        .AddNew
        txtfields(0) = HoldID$
        txtfields(5) = HoldCheck$
        If Trim(txtfields(5).Text) = "" Then
            cmdApply.Enabled = False
            cmdNSF.Enabled = False
            cmdPost.Enabled = False
            frNSF.Visible = False
        Else
            txtfields(3) = Format(Now, "mm/dd/yyyy")
            txtfields(6) = "Payment" '<<<-----  take a look at this, if invoice --"Payment" but if credit memo -- error in accounting logic
    '        TotalDue = grdDataGrid.Columns(7)
            TotalApplied = "$0.00"
            txtfields(12) = "$0.00"
            txtfields(4) = "$0.00"
            txtfields(2) = ""
            txtfields(4).Enabled = True
            cmdApply.Enabled = True
        End If
       '.Update
       '.Requery
       '.MoveLast
       'txtID = ![AR PAY ID]
    Else
        'If WhichHit = True Then
          'adoPrimaryRS.Bookmark = tb.Bookmark
        '  MsgBox "Please use the lookup... the search function is not added yet", vbInformation, "Information"
        'End If
        For Each Ctrl In Me.Controls
           If TypeOf Ctrl Is TextBox Or TypeOf Ctrl Is CheckBox Then
              Set Ctrl.DataSource = ADOprimaryrs
                
              If TypeOf Ctrl Is TextBox And Ctrl.DataField <> "" Then
                 If ADOprimaryrs("" & Ctrl.DataField & "").Type = 202 Then Ctrl.MaxLength = ADOprimaryrs("" & Ctrl.DataField & "").DefinedSize
              End If
           End If
        Next
        CheckAllData = True
        cmdbankAccount.Enabled = False
        txtfields(4).Locked = True
        'Rebuild Cash Receipts Work Table
        'Call RebuildTable
    End If
  End With
  
   'picButtons.Enabled = True
   'picStatBox.Enabled = True
   
  'Rebuild Cash Receipts Work Table
  'Call RebuildTable
  GetTextColor Me

  Exit Function
AR_PAY_Check_No_AfterUpdate_Error:
  Call LogError("Cash Receipts", "CheckAllData", Now, Err.Number, Err.Description, True)
  Resume Next

End Function

Private Function ValidateData() As Boolean

  If txtfields(5) = "" Then Exit Function
  If Len(Trim(txtfields(5))) = 0 Then Exit Function

  If txtfields(0) = "" Then
    MsgBox "Enter a customer first!", , "Error"
    ValidateData = True
    Exit Function
  End If

  If Len(Trim(txtfields(0))) = 0 Then
    MsgBox "Enter a customer first!", , "Error"
    ValidateData = True
    Exit Function
  End If
  
  Exit Function
AR_PAY_Check_No_BeforeUpdate_Error:
  Call LogError("Cash Receipts", "AR_PAY_Check_No_BeforeUpdate", Now, Err, Error, True)
  Resume Next

End Function

Private Sub grdDataGridSource(SQLstatement As String)

  Set adoPrimaryRS2 = New Recordset

    adoPrimaryRS2.Open SQLstatement, db, adOpenStatic, adLockOptimistic
        Set grdDatagrid.DataSource = Nothing
        grdDatagrid.HoldFields
        grdDatagrid.ReBind
        grdDatagrid.Refresh
    Set grdDatagrid.DataSource = adoPrimaryRS2
    If adoPrimaryRS2.RecordCount > 0 Then
       adoPrimaryRS2.MoveFirst
       Dim AppliedAmt As Currency
       Dim Balance As Currency
       Balance = 0
       Do While Not adoPrimaryRS2.EOF
          Balance = Balance + adoPrimaryRS2![Balance]
          AppliedAmt = AppliedAmt + adoPrimaryRS2![Applied Amount]
          adoPrimaryRS2.MoveNext
       Loop
          TotalDue = Format(Balance, "$###,####,##0.00")
          TotalApplied = Format(AppliedAmt, "$###,####,##0.00")
    End If
    
    Screen.MousePointer = vbNormal
    'dbase.Close
    'Set dbase = Nothing
End Sub

Private Sub cmdDel_Click()
Dim TempStr As String

Screen.MousePointer = vbHourglass

If ADOprimaryrs.EditMode = adEditAdd Then
    TempStr = txtfields(0).Text
    ADOprimaryrs.CancelUpdate
Else

If Datavalidate(True) = False Then
    Exit Sub
End If

    TempStr = txtfields(0).Text
    
    Dim Ctrl As Control
        For Each Ctrl In Me.Controls
           If TypeOf Ctrl Is TextBox Or TypeOf Ctrl Is CheckBox Then
              Set Ctrl.DataSource = Nothing
           End If
        Next
    Set grdDatagrid.DataSource = Nothing

  db.BeginTrans
    db.Execute "DELETE * FROM [AR Payment Invoice Cross Reference] where [AR CROSS Payment ID] = " & txtfields(2).Text
    db.Execute "DELETE * FROM [Cash Receipts Work]"
    db.Execute "DELETE * FROM [AR Payment Header] where [AR PAY ID] = " & txtfields(2).Text
  db.CommitTrans
End If
  txtfields(0).Text = TempStr
  txtfields(5).Text = ""
  PrevcheckNo = ""
  CallRebuildTable
  cmdDel.Enabled = False
Screen.MousePointer = vbNormal
End Sub


Private Sub cmdPost_Click()
  'On Error GoTo cmdPost_Click_Error
   
  Dim Success%

  'Force record save
  'DoCmd.RunMacro "Save Record"

  Screen.MousePointer = vbHourglass

  Call RefreshData

  db.BeginTrans
    Success% = AR_PostPayments()
    If Success% = False Then
      db.RollbackTrans
      MsgBox "Transaction NOT Posted."
    Else
      db.CommitTrans
      MsgBox "Transaction Posted."
      'chkFields(9) = 1
      ADOprimaryrs![AR PAY Posted YN] = True
      ADOprimaryrs![AR PAY Status] = "Posted"
      ADOprimaryrs.UpdateBatch adAffectAll
      'Call Form_Current
      'DoCmd.GoToControl "AR PAY Customer No"
      Call RebuildTable
    End If
    
  LockControl
  Screen.MousePointer = vbNormal
  
  Exit Sub
  
RecordLocked:
  db.RollbackTrans
  Exit Sub

UnableToPost:
  db.RollbackTrans
  Exit Sub

cmdPost_Click_Error:
  Call LogError("Cash Receipts", "cmdPost_Click", Now, Err, Error, True)
  Resume Next
End Sub

Private Function Datavalidate(MsgAppear As Boolean, Optional Str1 As String, Optional str2 As String) As Boolean
Dim i As Integer

For i = txtfields.LBound To txtfields.UBound
    Select Case i
    Case 0, 1, 2, 4, 5
        If txtfields(i) = "" Then
          If MsgAppear = True Then MsgBox Str1 & " " & lblLabels(i) & " " & str2, , "Error"
          Datavalidate = False
          Exit Function
        End If
    End Select
Next

For i = chkFields.LBound To chkFields.UBound
Select Case i
Case 7, 8, 9, 10, 13
    If chkFields(i).Value = 1 Then
        MsgBox "This is a " & Left(chkFields(i).Caption, Len(chkFields(i).Caption) - 1) & " transaction. Your request is denied", vbInformation, "Information"
        Datavalidate = False
        Exit Function
    End If
End Select
Next

Datavalidate = True
  
  Exit Function
DataValidate_Error:
  Call LogError("Cash Receipts", "DataValidate", Now, Err.Number, Err.Description, True)
  Resume Next

End Function


Private Sub cmdbankAccount_Click()
    Dim SQLstatement As String
    Dim No As Integer
    Dim ghead As String
    Dim fhead As String

    No = 1216
    SQLstatement = "select [BANK ACCT ID], [BANK ACCT Name]" & _
                    "from [BANK Accounts]"
    ghead = "Account Description"
    fhead = "Account No//Description"
    
    AllLookup.GetWhichTable No, SQLstatement, ghead, fhead
    'AllLookup.Show vbModal
    If txtfields(1) <> "" Then
        txtfields(1).SetFocus
        txtfields(7) = NZ(DLookup("[GL COA Account Balance]", "[GL Chart Of Accounts]", "[GL COA Account No] = '" & txtfields(1).Text & "'"))
        txtfields(7) = Format(txtfields(7), "$###,###,###,##0.00")
    End If
End Sub

Private Sub cmdCheckNo_Click()
'Dim tempStr As String

   'tempStr = txtFields(5).Text
   AllLookup.GetWhichTable 1215, "select [AR PAY Check No],[AR PAY Customer No],[AR PAY Type],[AR PAY Transaction Date],[AR PAY Amount]" & _
   "from [AR Payment Header] where [AR PAY Customer No]='" & txtfields(0) & "'", "Customer Selection", _
   "Check No//Customer ID//Payment type//Transaction Date//Amount"
   
   'AllLookup.Show vbModal
   Screen.MousePointer = vbHourglass
   'txtFields(5).SetFocus
   'cmdCheckNo.SetFocus
   
   If PrevcheckNo = txtfields(5).Text Then
        Screen.MousePointer = vbNormal
        Exit Sub
   End If
    PrevcheckNo = txtfields(5).Text
    CallRebuildTable
    LockControl
    If txtfields(1) <> "" Then
        txtfields(1).SetFocus
        txtfields(7) = NZ(DLookup("[GL COA Account Balance]", "[GL Chart Of Accounts]", "[GL COA Account No] = '" & txtfields(1).Text & "'"))
        txtfields(7) = Format(txtfields(7), "$###,###,###,##0.00")
    End If
    Screen.MousePointer = vbNormal
End Sub

Private Sub cmdNSF_Click()
   Screen.MousePointer = vbHourglass
   
  'On Error GoTo cmdNSF_Click_Error
   
  Dim Success%

  'Is this a credit memo
  If txtfields(6) = "Credit Memo" Then
    MsgBox "Can't mark a credit memo NSF!", , "Error"
    Screen.MousePointer = vbNormal
    Exit Sub
  End If
  If chkFields(9).Value <> 1 Then
    MsgBox "This transaction is not posted yet. Please use delete button", , "Error"
    Screen.MousePointer = vbNormal
    Exit Sub
  End If
  
  Dim Response%
  Response% = MsgBox("Mark this payment NSF?", vbYesNo, "NSF")
  If Response% = vbNo Then
    Screen.MousePointer = vbNormal
    Exit Sub
  End If

  'Force record save
  'DoCmd.RunMacro "Save Record"

  Call RefreshData

  db.BeginTrans
    Success% = PostNSF()
    If Success% = False Then
      db.RollbackTrans
      MsgBox "Transaction NOT marked NSF."
    Else
      db.CommitTrans
      MsgBox "Transaction marked NSF."
      'chkFields(13) = True
      ADOprimaryrs![AR PAY NSF] = True
      ADOprimaryrs![AR PAY Status] = "Non Sufficient Fund"
      ADOprimaryrs.UpdateBatch adAffectAll
      'Form_Current
      Call RebuildTable
    End If
    
  LockControl
  Screen.MousePointer = vbNormal
  
  Exit Sub
  
NSFLocked:
  db.RollbackTrans
  Exit Sub

UnableToMark:
  db.RollbackTrans
  Exit Sub

cmdNSF_Click_Error:
  Call LogError("Cash Receipts", "cmdNSF_Click", Now, Err, Error, True)
  Resume Next

End Sub

Private Sub cmdApply_Click()
   Screen.MousePointer = vbHourglass
  'On Error GoTo cmdApply_Click_Error
  
  If cmdApply.Caption = "&Save" Then
    'db.BeginTrans
    'ADOprimaryrs![AR PAY UnApplied Amount] = TotalApplied
    ADOprimaryrs.UpdateBatch adAffectAll
    ADOprimaryrs.Requery
    txtfields(2) = ADOprimaryrs![AR PAY ID]
    txtfields(12).Text = "$0.00"
    cmdApply.Caption = "&Apply"
    cmdPost.Enabled = True
    cmdNSF.Enabled = True
    cmdApply.Enabled = False
    Screen.MousePointer = vbNormal
    Exit Sub
  End If
Dim i As Integer

For i = txtfields.LBound To txtfields.UBound
    Select Case i
    Case 0, 1, 4, 5
      If i = 4 Then
        If txtfields(i) = "$0.00" Then
          MsgBox "Transaction cannot be made with $0.00", vbInformation, "Information"
          Screen.MousePointer = vbNormal
          Exit Sub
        End If
      Else
        If txtfields(i) = "" Then
          MsgBox "There is an empty data in " & " " & lblLabels(i), , "Information"
          Screen.MousePointer = vbNormal
          Exit Sub
        End If
      End If
    End Select
Next
  
  cmdUpdate_Click
  
  Dim Response%
  Response% = MsgBox("This will automatically distribute the check amount starting with the oldest listed invoice.", vbOKCancel, "Auto Apply")
  If Response% = vbCancel Then
    Screen.MousePointer = vbNormal
    Exit Sub
  End If

  Set grdDatagrid.DataSource = Nothing

  'Automatically apply payment to invoices
  Dim rsWork As ADODB.Recordset
  Dim AmountLeft#

  AmountLeft# = txtfields(4)
  Dim Balance#

  Set rsWork = New Recordset
  rsWork.Open "Select * from [Cash Receipts Work]", db, adOpenStatic, adLockOptimistic
  'On Error Resume Next
  rsWork.MoveFirst
  If Err = 0 Then
    'On Error GoTo cmdApply_Click_Error
    Do While Not rsWork.EOF
      Balance# = rsWork("Balance") + rsWork("Applied Amount")
      If Balance# <= AmountLeft# Then
          rsWork("Balance") = 0
          rsWork("Applied Amount") = Balance#
        rsWork.Update
        AmountLeft# = AmountLeft# - Balance#
      Else
          rsWork("Balance") = Balance# - AmountLeft#
          rsWork("Applied Amount") = AmountLeft#
        rsWork.Update
        AmountLeft# = 0
      End If
      If AmountLeft# <= 0 Then Exit Do
      rsWork.MoveNext
    Loop
  End If

  rsWork.Close
  'Forms![Cash Receipts].[Cash Receipts Detail].Form.Requery
  
  grdDataGridSource "select *  from [Cash Receipts Work] Order by [Reference]"
  
  txtfields(12).Text = "$0.00"
  cmdApply.Caption = "&Save"
  txtfields(4).Locked = True
  
  LockControl
  GetTextColor Me
Screen.MousePointer = vbNormal
  Exit Sub
cmdApply_Click_Error:
  Call LogError("Cash Receipts", "cmdApply_Click", Now, Err, Error, True)
  Resume Next
    
End Sub

Private Sub RedoNumbers()

  'On Error GoTo RedoNumbers_Error

  Dim rs As ADODB.Recordset
  Dim rs2 As ADODB.Recordset

  'xxx 1/7/97  7.2b
  db.BeginTrans

  db.Execute "DELETE * FROM [Receipt Numbers]"

  Set rs = New Recordset
  rs.Open "[AR Payment Header]", db, adOpenStatic, adLockOptimistic, adcmdtablle
  Set rs2 = New Recordset
  rs2.Open "[Receipt Numbers]", db, adOpenStatic, adLockOptimistic, adcmdtablle

  'On Error Resume Next

  rs.MoveFirst
  Do While Not rs.EOF
    rs2.AddNew
      rs2("Check No") = rs("AR PAY Check No")
      rs2("Bank") = rs("AR PAY Bank Account")
    rs2.Update
    rs.MoveNext
  Loop
  
  'xxx 1/7/97  7.2b
  db.CommitTrans

  Exit Sub
RedoNumbers_Error:
  Call LogError("Cash Receipts", "RedoNumbers", Now, Err, Error, True)
  Resume Next

End Sub


Private Sub RefreshData()
'maybe we could join this with PostData Module == almost alike
  'On Error GoTo RefreshData_Error

  Dim rsWork As ADODB.Recordset
  'Dim rsCross As ADODB.Recordset
  Dim rsSales As ADODB.Recordset
  Dim SaleID$

  If txtfields(2) = 0 Or txtfields(2) = "" Then Exit Sub

  Set rsWork = New Recordset
  rsWork.Open "[Cash Receipts Work]", db, adOpenStatic, adLockOptimistic, adCmdTable
  'Set rsCross = New Recordset
  'rsCross.Open "[AR Payment Invoice Cross Reference]", db, adOpenStatic, adLockOptimistic, adCmdTable
  Set rsSales = New Recordset
  rsSales.Open "[AR Sales]", db, adOpenStatic, adLockOptimistic, adCmdTable

  'rsSales.Index = "Ext Document #"

  'Scratch detail
  db.Execute "DELETE * FROM [AR Payment Invoice Cross Reference] where [AR CROSS Payment ID] = " & txtfields(2).Text, , adCmdText

  'Save detail information
  Dim TotalAppliedCurr#
  TotalAppliedCurr# = 0

  'On Error Resume Next
  If rsWork.RecordCount > 0 Then
  rsWork.MoveFirst
  'If Err = 0 Then
    'On Error GoTo RefreshData_Error
    Do While Not rsWork.EOF
      SaleID$ = rsWork("Reference")
      rsSales.MoveFirst
      rsSales.Find "[AR SALE Ext Document #]='" & SaleID$ & "'"
      If rsSales.EOF Then
        'Should not happen
      Else
        'Update Sales Record
        If chkFields(9) = 1 And chkFields(13) = 0 And chkFields(7) = 0 Then
          'rsSales.Edit
            rsSales("AR SALE Amount Paid") = rsSales("AR SALE Total") - rsWork("Balance")
            rsSales("AR SALE Balance Due") = rsWork("Balance")
          rsSales.Update
        End If
        If rsWork("Applied Amount") > 0 Then  'rsWork("Discount") > 0 Or rsWork("Write Off") > 0 Or rsWork("Applied Amount") > 0 Then
            'Write a cross reference record
            SQLstatement = "INSERT INTO [AR Payment Invoice Cross Reference]"
            SQLstatement = SQLstatement & " ([AR CROSS Payment ID],[AR CROSS Payed ID],[AR CROSS Discount Taken],"
            SQLstatement = SQLstatement & "[AR CROSS Write Off Amount],[AR CROSS Applied Amount],[AR CROSS Cleared])"
            SQLstatement = SQLstatement & " VALUES (" & txtfields(2).Text & "," & rsSales("AR SALE Document #") & ","
            SQLstatement = SQLstatement & rsWork("Discount") & "," & rsWork("Write Off") & "," & rsWork("Applied Amount") & ",False)"
            'Debug.Print SQLstatement
            
            db.Execute SQLstatement
          'rsCross.AddNew
          '  rsCross("AR CROSS Payment ID") = txtFields(2)
          '  rsCross("AR CROSS Payed ID") = rsSales("AR SALE Document #")
          '  rsCross("AR CROSS Discount Taken") = rsWork("Discount")
          '  rsCross("AR CROSS Write Off Amount") = rsWork("Write Off")
          '  rsCross("AR CROSS Applied Amount") = rsWork("Applied Amount")
            TotalAppliedCurr# = TotalAppliedCurr# + rsWork("Applied Amount")
          '  rsCross("AR CROSS Cleared") = False
          'rsCross.Update
        Else
        End If
      End If
      rsWork.MoveNext
    Loop
  End If


  TotalAppliedCurr# = txtfields(4).Text - TotalAppliedCurr#
  txtfields(12).Text = Format(TotalAppliedCurr#, "$###,###,##0.00")
  Exit Sub
RefreshData_Error:
  Call LogError("Cash Receipts", "RefreshData", Now, Err, Error, True)
  Resume Next

End Sub

Private Sub RebuildTable()

  'Rebuild Cash Receipts Table
  
  'On Error GoTo RebuildTable_Error
  
  db.Execute "DELETE * FROM [Cash Receipts Work]"
  
  'If txtfields(0).Text = "" Or Len(Trim(txtfields(0).Text)) = 0 Then
  '  adoPrimaryRS.Requery
  '  Exit Sub
  'End If

  Dim NoCheck%
  NoCheck% = False

    If txtfields(5) = "" Or Len(Trim(txtfields(5))) = 0 Then
        NoCheck% = True
    End If
  
    If ADOprimaryrs Is Nothing Then
    Else
        If ADOprimaryrs.RecordCount = 1 And ADOprimaryrs.EditMode = adEditAdd Then NoCheck% = True
    End If
  
  Screen.MousePointer = vbHourglass
  
  'xxx 4/17/97 7.3
  'Dim rsCashReceipts As ADODB.Recordset
  
  If NoCheck% = False Then
    'See if this check is already in the system
    'Set rsCashReceipts = New Recordset
    Dim FoundCashReceipt As Boolean
    FoundCashReceipt = CheckDocument("SELECT * FROM [AR Payment Header] WHERE [AR PAY Customer No]='" & txtfields(0) & "' AND [AR PAY Check No]='" & txtfields(5) & "'", True)
    'rsCashReceipts.Index = "PrimaryKey"
    'rsCashReceipts.Find "[AR PAY Customer No]='" & txtfields(0) & "' AND [AR PAY Check No]='" & txtfields(5) & "'"
    'rsCashReceipts.Find "[AR PAY Check No]='" & txtfields(0) & "'"
    If FoundCashReceipt = True Then
      NoCheck% = True
    Else
      NoCheck% = False
    End If
  End If
    
  Dim rsSales As ADODB.Recordset
  Dim rsWork As ADODB.Recordset
  Dim rsCross As ADODB.Recordset
  Dim SaleID&
  If NoCheck% = True Then
    'Load all invoices for this customer w/ a balance
    Set rsSales = New Recordset
    'rsSales.Open "SELECT * FROM [AR Sales] where [AR SALE Posted YN] = True and [AR SALE Balance Due] >= .01 and [AR SALE Customer ID] = '" & txtfields(0) & "' and [AR SALE Check Number]='" & txtfields(5) & "' AND [AR SALE Document Type] in ('Invoice','Sales Memo','Beginning Balance','Finance Charge')", db, adOpenStatic, adLockOptimistic
    rsSales.Open "SELECT * FROM [AR Sales] where [AR SALE Posted YN] = True and [AR SALE Balance Due] >= .01 and [AR SALE Customer ID] = '" & txtfields(0) & "' AND [AR SALE Document Type] in ('Invoice','Sales Memo','Beginning Balance','Finance Charge')", db, adOpenForwardOnly, adLockOptimistic
    
    Set rsWork = New Recordset
    rsWork.Open "[Cash Receipts Work]", db, adOpenForwardOnly, adLockOptimistic, adCmdTable
    'On Error Resume Next
    
    If rsSales.RecordCount = 0 Then
        MsgBox "No Transaction on this company.", vbInformation, "Information"
        TotalDue = ""
        TotalApplied = ""
        cmdCheckNo.Enabled = False
        picButtons.Enabled = False
        picStatBox.Enabled = False
        For Each Ctrl In Me.Controls
           If TypeOf Ctrl Is TextBox Or TypeOf Ctrl Is CheckBox Then
              Set Ctrl.DataSource = Nothing
              If TypeOf Ctrl Is TextBox Then
                Ctrl.Text = ""
              Else
                Ctrl.Value = 0
              End If
           End If
        Next
        Screen.MousePointer = vbNormal
        'Exit Sub
    Exit Sub
    End If
    cmdCheckNo.Enabled = True
    rsSales.MoveFirst
    'Exit Sub
    'If Err = 0 Then
      'On Error Resume Next
      'Load these records into work table
      Do While Not rsSales.EOF
        rsWork.AddNew
          rsWork("Reference") = CStr(rsSales("AR SALE Ext Document #"))
          rsWork("Date") = rsSales("AR SALE Date")
          rsWork("Original Amount") = rsSales("AR SALE Total")
          'Calculate discount amount
          rsWork("Amount Paid") = rsSales("AR SALE Amount Paid")
          If rsWork("Amount Paid") > 0 Then
            rsWork("Discount") = 0
          Else
            rsWork("Discount") = GetARInvoiceDiscount(CLng(rsSales("AR SALE Document #")), txtfields(3), False)
          End If
          rsWork("Write Off") = 0
          rsWork("Applied Amount") = 0
          rsWork("Balance") = rsSales("AR SALE Balance Due") - rsWork("Discount")
        rsWork.Update
        rsSales.MoveNext
      Loop
    'End If
        
  Else
    'Load all invoices and invoices paid w/ this check
    Set rsSales = New Recordset
    'rsSales.Open "SELECT * FROM [AR Sales] where [AR SALE Posted YN] = True and [AR SALE Balance Due] >= .01 and [AR SALE Customer ID] = '" & txtfields(0) & "' and [AR SALE Check Number]='" & txtfields(5) & "' AND [AR SALE Document Type] in ('Invoice','Sales Memo','Beginning Balance','Finance Charge')", db, adOpenStatic, adLockOptimistic
    rsSales.Open "SELECT * FROM [AR Sales] where [AR SALE Posted YN] = True and [AR SALE Customer ID] = '" & txtfields(0) & "' AND [AR SALE Document Type] in ('Invoice','Sales Memo','Beginning Balance','Finance Charge')", db, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    Set rsCross = New Recordset
    rsCross.Open "[AR Payment Invoice Cross Reference]", db, adOpenForwardOnly, adLockOptimistic, adCmdTable
    'rsCross.Index = "PaymentPayed"
    
    Set rsWork = New Recordset
    rsWork.Open "[Cash Receipts Work]", db, adOpenForwardOnly, adLockOptimistic, adCmdTable
    'On Error Resume Next
    If rsSales.RecordCount = 0 Then
        Screen.MousePointer = vbNormal
        Exit Sub
    End If
    'If Err = 0 Then
      'On Error GoTo RebuildTable_Error
      'Load these records into work table
      Do While Not rsSales.EOF '<<<-----------------------------------------
        SaleID& = rsSales("AR SALE Document #")
        If rsSales("AR SALE Balance Due") = 0 Then
          'rsCross.seek "[AP CROSS ID]=", Forms![Cash Receipts]![AR PAY ID], SaleID&  'AR CROSS Payed ID
          rsCross.MoveFirst
          rsCross.Find "[AR CROSS Payed ID]=" & SaleID&  'AR CROSS Payed ID
          If rsCross.EOF Then
            GoTo SkipIt
          End If
        End If
        rsWork.AddNew
          rsWork("Reference") = CStr(rsSales("AR SALE Ext Document #"))
          rsWork("Date") = rsSales("AR SALE Date")
          rsWork("Original Amount") = rsSales("AR SALE Total")
          'rsCross.Seek "=", Forms![Cash Receipts]![AR PAY ID], SaleID&
          'If rsCross.EOF Then'<<<-----------------------------------------
          rsCross.MoveFirst
          rsCross.Find "[AR CROSS Payed ID]=" & SaleID&  'AR CROSS Payed ID
          If rsCross.EOF Then
            rsWork("Amount Paid") = rsSales("AR SALE Amount Paid")
            'xxx 12/2/96 7.2a
           ' If rsWork("Amount Paid") > 0 Then
              rsWork("Discount") = 0
           ' Else
           '   rsWork("Discount") = GetARInvoiceDiscount(CLng(rsSales("AR SALE Document #")), Me![AR PAY Transaction Date], False)
           ' End If
            rsWork("Write Off") = 0
            rsWork("Applied Amount") = 0
            rsWork("Balance") = rsSales("AR SALE Balance Due") - rsWork("Discount")
          Else
            rsWork("Discount") = rsCross("AR CROSS Discount Taken")
            rsWork("Write Off") = rsCross("AR CROSS Write Off Amount")
            rsWork("Applied Amount") = rsCross("AR CROSS Applied Amount")
            If chkFields(9).Value = 1 Then
              rsWork("Amount Paid") = rsSales("AR SALE Amount Paid") - rsCross("AR CROSS Applied Amount") - rsCross("AR CROSS Discount Taken") - rsCross("AR CROSS Write Off Amount")
              rsWork("Balance") = rsSales("AR SALE Balance Due")
            Else
              rsWork("Amount Paid") = rsSales("AR SALE Amount Paid")
              rsWork("Balance") = rsSales("AR SALE Balance Due") - rsCross("AR CROSS Applied Amount") - rsCross("AR CROSS Discount Taken") - rsCross("AR CROSS Write Off Amount")
            End If
          End If
        rsWork.Update
SkipIt:
        rsSales.MoveNext
      Loop
    End If
        
  'End If

  'Refresh the tables
  rsWork.Close

  'adoPrimaryRS.Requery
  'Me![Cash Receipts Detail].Form![Total Due].Requery

  'DoCmd.Hourglass False
  Screen.MousePointer = vbNormal

  Exit Sub
RebuildTable_Error:
  Call LogError("Cash Receipts", "RebuildTable", Now, Err, Error, True)
  Resume Next

End Sub


Private Sub cmdCashCustomer_Click()
   Dim PreviousCustId As String
   Dim Ctrl As Control
   PreviousCustId = txtfields(0)
   AllLookup.GetWhichTable 1210, "Select [AR CUST Customer ID],[AR CUST Name],[AR CUST Address 1]," & _
   "[AR CUST Address 2],[AR CUST City],[AR CUST State],[AR CUST Postal],[AR CUST Country] From " & _
   "[AR Customer] ", "Customer Particular", _
   "Customer ID//Customer Name//Address 1//Address 2//City//State//Postal//Country"
   
   'AllLookup.Show vbModal
   If PreviousCustId = txtfields(0) Then Exit Sub
   
Screen.MousePointer = vbHourglass
'   Set adoPrimaryRS = New Recordset
'   adoPrimaryRS.Open "select [AR PAY ID],[AR PAY Type],[AR PAY Check No],[AR PAY Customer No],[AR PAY Transaction Date],[AR PAY Amount],[AR PAY Bank Account],[AR PAY Posted YN],[AR PAY Reconciled],[AR PAY Cleared],[AR PAY Deposited YN],[AR PAY NSF],[AR PAY UnApplied Amount],[AR PAY Status],[AR PAY Notes] from [AR Payment Header] WHERE [AR PAY Customer No]='" & txtTemp & "'", db, adOpenStatic, adLockOptimistic
'   txtfields(0) = txtTemp
   txtfields(5) = ""
   RebuildTable
   
   cmdPost.Enabled = False
   cmdNSF.Enabled = False
   cmdApply.Enabled = False
   lblTemp = ""
   'adoPrimaryRS.Open "select * from [AR Payment Header] WHERE [AR PAY Customer No]='" & txtfields(0) & "'", db, adOpenStatic, adLockOptimistic
'   If adoPrimaryRS.RecordCount = 0 Then
'      MsgBox "No Transaction on this company yet.", vbInformation, "Information"
'      cmdCheckNo.Enabled = False
'      picButtons.Enabled = False
'      picStatBox.Enabled = False
'      lblcashReceiptsTrue.Caption = ""
        For Each Ctrl In Me.Controls
           If TypeOf Ctrl Is TextBox Or TypeOf Ctrl Is CheckBox Then
              Set Ctrl.DataSource = Nothing
                If TypeOf Ctrl Is TextBox Then
                    If Ctrl.Text <> txtfields(0).Text Then Ctrl.Text = ""
                ElseIf TypeOf Ctrl Is CheckBox Then
                    Ctrl.Value = 0
                End If
           End If
        Next
        'grdDataGrid.HoldFields
        'Set grdDataGrid.DataSource = Nothing
        'grdDataGrid.ReBind
        'grdDataGrid.Refresh
        grdDataGridSource "select *  from [Cash Receipts Work] Order by [Reference]"
'      Exit Sub
'   End If
'   lblcashReceiptsTrue.Caption = lblcashReceipts.Caption
'   lblcashReceiptsTrue.Visible = True
   'lblcashReceipts.Visible = True
'   cmdCheckNo.Enabled = True
'   picButtons.Enabled = True
'   picStatBox.Enabled = True
'   For Each ctrl In Me.Controls
'      If TypeOf ctrl Is TextBox Or TypeOf ctrl Is CheckBox Then
'         Set ctrl.DataSource = adoPrimaryRS
'      End If
'   Next
'   txtFields_LostFocus 5
Screen.MousePointer = vbNormal
End Sub


Private Sub cmdPrint_Click()
If txtfields(0) = "" Or txtfields(5) = "" Then
    MsgBox "No check to print", vbInformation, "Information"
    Exit Sub
End If

    If ADOprimaryrs Is Nothing Then
    Else
        If ADOprimaryrs.EditMode = adEditAdd Then
            MsgBox "New check is in progress.", vbInformation, "Information"
        End If
    End If
End Sub

Private Sub Form_Load()
Screen.MousePointer = vbHourglass
  Set db = New Connection
  db.CursorLocation = adUseClient
  db.Open gblADOProvider
  
  lblcashReceiptsTrue = ""
  frNSF.ZOrder 0
  GetTextColor Me
  mbDataChanged = False
Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Resize()
  'On Error Resume Next
  'This will resize the grid when the form is resized
  'grdDataGrid.Width = frPrimary.Width
  'grdDataGrid.Height = frPrimary.Height - grdDataGrid.Top - 30 - picButtons.Height - picStatBox.Height
  If fMainForm.WindowState = 1 Then Exit Sub
  If Me.WindowState = 0 Then
  ElseIf Me.WindowState = 2 Then
    GoTo SkipResize
  Else
    Exit Sub
  End If
  
  Me.Width = 10335
  Me.Height = 7620
SkipResize:
  frPrimary.Left = (Me.ScaleWidth - frPrimary.Width) / 2
  Label1(2).Left = frPrimary.Left
  frPrimary.Top = (Me.ScaleHeight - frPrimary.Height) / 2 + 230
  
  lblStatus.Width = Me.Width - 1500
  cmdNext.Left = lblStatus.Width + 700
  cmdLast.Left = cmdNext.Left + 340
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If mbEditFlag Or mbAddNewFlag Then Exit Sub

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

    'updates the checklist Customers
  Screen.MousePointer = vbHourglass
      Set ADOprimaryrs = Nothing
      db.Close
      Set db = Nothing
  Screen.MousePointer = vbDefault
  Set frm_AR_Cash_Receipts = Nothing
Exit Sub
FormErr:
  MsgBox Err.Description
  Screen.MousePointer = vbDefault
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
    SetButtons False
  End With

  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
     If Not DataDelete(ADOprimaryrs, Me) Then
     End If
End Sub

Private Sub cmdRefresh_Click()
    RefreshButton ADOprimaryrs, grdDatagrid
End Sub

Private Sub cmdEdit_Click()
  'On Error GoTo EditErr

  lblStatus.Caption = "Edit record"
  mbEditFlag = True
  SetButtons False
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
Dim FlagStatus As Boolean
    
  FlagStatus = False

  Call UpdateButton(ADOprimaryrs, FlagStatus, mbAddNewFlag)
  
  mbEditFlag = Not FlagStatus
  mbAddNewFlag = Not FlagStatus
  SetButtons FlagStatus
  mbDataChanged = Not FlagStatus
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
'  cmdCancel.Visible = Not bVal
  cmdDelete.Visible = bVal
  cmdClose.Visible = bVal
  cmdRefresh.Visible = bVal
  cmdNext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdLast.Enabled = bVal
  cmdPrevious.Enabled = bVal
End Sub

Private Sub txtfields_GotFocus(Index As Integer)
    TxtGotFocus txtfields(Index)
End Sub

Private Sub txtFields_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
Case 4
    keyResponse = CtrlValidate(KeyAscii, "0123456789.")
    If keyResponse = True Then
    Else
       KeyAscii = 0
    End If
Case 5
    keyResponse = CtrlValidate(KeyAscii, "0123456789")
    If keyResponse = True Then
    Else
       KeyAscii = 0
    End If
End Select
End Sub

Private Sub txtFields_LostFocus(Index As Integer)
Select Case Index
Case 5
     If PrevcheckNo = txtfields(5).Text Then Exit Sub
     txtfields(5).Text = Trim(txtfields(5).Text)
     PrevcheckNo = txtfields(5).Text
     CallRebuildTable
End Select
  GetTextColor Me
End Sub
