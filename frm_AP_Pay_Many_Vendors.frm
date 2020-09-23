VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_AP_Pay_Many_Vendors 
   Caption         =   "Multiple Vendor Payment"
   ClientHeight    =   7215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9375
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7215
   ScaleWidth      =   9375
   Begin VB.Frame frPrimary 
      Height          =   6735
      Left            =   0
      TabIndex        =   12
      Top             =   480
      Width           =   9375
      Begin VB.TextBox txtFields 
         DataField       =   " "
         Height          =   285
         Index           =   1
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtFields 
         DataField       =   " "
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
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox txtFields 
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
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtFields 
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
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox txtFields 
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
         Index           =   14
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox txtFields 
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
         Left            =   7800
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton cmdbankAccount 
         Height          =   285
         Left            =   3600
         Picture         =   "frm_AP_Pay_Many_Vendors.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton cmdDate 
         Height          =   285
         Index           =   1
         Left            =   1800
         Picture         =   "frm_AP_Pay_Many_Vendors.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton cmdDate 
         Height          =   285
         Index           =   2
         Left            =   1440
         Picture         =   "frm_AP_Pay_Many_Vendors.frx":0724
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1560
         Width           =   375
      End
      Begin VB.CommandButton cmdPost 
         Caption         =   "&Post"
         Height          =   855
         Left            =   8280
         Picture         =   "frm_AP_Pay_Many_Vendors.frx":0CFE
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton cmdDate 
         Caption         =   "All"
         Height          =   285
         Index           =   0
         Left            =   1800
         Picture         =   "frm_AP_Pay_Many_Vendors.frx":1140
         TabIndex        =   16
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox txtFields 
         DataField       =   " "
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
         Index           =   0
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtFields 
         DataField       =   " "
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
         Index           =   2
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox txtFields 
         DataField       =   " "
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
         Index           =   8
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   720
         Width           =   1575
      End
      Begin VB.Frame Frame1 
         Height          =   4695
         Left            =   120
         TabIndex        =   27
         Top             =   1920
         Width           =   9135
         Begin MSDataGridLib.DataGrid grdDataGrid 
            Height          =   2715
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   8880
            _ExtentX        =   15663
            _ExtentY        =   4789
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
            Caption         =   "Payment Transaction"
            ColumnCount     =   6
            BeginProperty Column00 
               DataField       =   "PH Selected"
               Caption         =   "Pay"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   5
                  Format          =   ""
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
            BeginProperty Column01 
               DataField       =   "PH Vendor ID"
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
            BeginProperty Column02 
               DataField       =   "PH Vendor Name"
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
            BeginProperty Column03 
               DataField       =   "PH Total Owed"
               Caption         =   "Owed"
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
               DataField       =   "PH Payment Total"
               Caption         =   "Payment"
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
            BeginProperty Column05 
               DataField       =   "PH Due Date"
               Caption         =   "Latest Due Date"
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
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
                  ColumnWidth     =   599.811
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1305.071
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   2294.929
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   1395.213
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   1395.213
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   1260.284
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid DataGridDetaila 
            Height          =   1425
            Left            =   120
            TabIndex        =   29
            Top             =   3120
            Visible         =   0   'False
            Width           =   8880
            _ExtentX        =   15663
            _ExtentY        =   2514
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   -2147483624
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
            Caption         =   "Purchase Transaction"
            ColumnCount     =   8
            BeginProperty Column00 
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
            BeginProperty Column01 
               DataField       =   "AP Pay Type"
               Caption         =   "Transaction"
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
               DataField       =   "AP PO Due Date"
               Caption         =   "Due Date"
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
            BeginProperty Column03 
               DataField       =   "Original Amount"
               Caption         =   "Orig Amount"
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
               DataField       =   "Discount"
               Caption         =   "Disc. Amt"
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
            BeginProperty Column05 
               DataField       =   "Applied Amount"
               Caption         =   "Applied Amt"
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
               DataField       =   "Balance"
               Caption         =   "Balance"
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
               DataField       =   "Amount Paid"
               Caption         =   "Amount Paid"
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
               BeginProperty Column00 
                  ColumnWidth     =   1230.236
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1244.976
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   1154.835
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   1244.976
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   1140.095
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   1140.095
               EndProperty
               BeginProperty Column06 
                  ColumnWidth     =   1140.095
               EndProperty
               BeginProperty Column07 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1035.213
               EndProperty
            EndProperty
         End
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bank Account"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   39
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Check Total:"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   4200
         TabIndex        =   38
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Check Date"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   37
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Account Balance"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   2400
         TabIndex        =   36
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Due Date"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   35
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "End Balance:"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   12
         Left            =   7800
         TabIndex        =   34
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label tmdDate 
         Height          =   255
         Left            =   600
         TabIndex        =   33
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Applied Amout"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   6000
         TabIndex        =   32
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Unapplied Amount"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   6000
         TabIndex        =   31
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Check Number"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   16
         Left            =   4200
         TabIndex        =   30
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.PictureBox picStatBox 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   9375
      TabIndex        =   0
      Top             =   6915
      Visible         =   0   'False
      Width           =   9375
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Picture         =   "frm_AP_Pay_Many_Vendors.frx":171A
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Picture         =   "frm_AP_Pay_Many_Vendors.frx":1A5C
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Picture         =   "frm_AP_Pay_Many_Vendors.frx":1D9E
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Picture         =   "frm_AP_Pay_Many_Vendors.frx":20E0
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
   Begin VB.PictureBox picButtons 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   11160
      TabIndex        =   6
      Top             =   5640
      Visible         =   0   'False
      Width           =   11160
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Multiple Vendor Payment"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   40
      Top             =   120
      Width           =   7665
   End
End
Attribute VB_Name = "frm_AP_Pay_Many_Vendors"
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
'Dim mbDataChanged As Boolean

Dim db As ADODB.Connection
Dim strSQL As String
Dim TempRec As ADODB.Recordset
Dim NowLoad As Boolean
Dim NextCheck$
Dim TempStr As String


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
    
    AllLookup.GetWhichTable No, SQLstatement, ghead, fhead, db
    'AllLookup.Show vbModal
    BankAcct
    txtFields(8).SetFocus
End Sub

Private Sub BankAcct()

  ''On Error GoTo Bank_Acct_BeforeUpdate_Error

  If txtFields(1).Text = "" Then Exit Sub
    txtFields(5) = AcctBalance("balance", txtFields(1).Text, db)
    Call CalcManyTotals
    
    'get a valid checkNumber
    Dim GetCheckNo As String
    GetCheckNo = CheckNumberCHQ("read", db, txtFields(1).Text)
    If GetCheckNo <> "" Then
      txtFields(8).Text = GetCheckNo
    Else
      txtFields(8).Text = ""
    End If

  'If Len(Trim(Me![Bank Acct])) = 0 Then Exit Sub

  'Make sure this is a valid cash account
  'Dim rsGL As ADODB.Recordset
  'Set rsGL = New ADODB.Recordset
  'rsGL.Open "SELECT GL COA Account Balance FROM [GL Chart Of Accounts] WHERE [GL COA Account No]='" & txtFields(1) & "'", db, adOpenKeyset, adLockOptimistic
  'rsGL.Index = "PrimaryKey"
  'rsGL.Seek "=", Me![Bank Acct]
  'If rsGL.RecordCount = 0 Then
  '  MsgBox "Not a valid GL Account!", , "Error"
    'Cancel = True
  '  Exit Sub
  'End If

  'If rsGL("GL COA Asset Type") = "Cash" Then
    'txtFields(5) = FormatCurr(rsGL("GL COA Account Balance") )
    'Call CalcManyTotals
  'Else
  '  MsgBox "Not a valid Cash Account!", , "Error"
  '  Cancel = True
  '  Exit Sub
  'End If
    
  'Dim GetCheckNo As String
  'GetCheckNo = GetcheckNumber(db, txtFields(1))
  'If GetCheckNo <> "Error" Then
  '  txtFields(8) = GetCheckNo
  'Else
  'End If
  
  'Dim rsBank As ADODB.Recordset
  'Set rsBank = New ADODB.Recordset
  'rsBank.Open "SELECT * FROM [Bank Accounts] WHERE [BANK ACCT ID]='" & txtFields(1) & "'", db, adOpenStatic, adLockOptimistic
  'rsBank.Index = "PrimaryKey"
  'rsBank.Seek "=", Me![Bank Acct]
  'Me![Next Check No] = rsBank("BANK ACCT Next Check No")
  'If rsBank.RecordCount = 0 Then
  '   MsgBox "There is an error on Bank setup!", vbCritical, "Critical Error"
  '   Exit Sub
  'End If
  
  'txtFields(8) = rsBank("BANK ACCT Next Check No")

  Exit Sub
Bank_Acct_BeforeUpdate_Error:
  Call ErrorLog("Pay Many", "Bank_Acct_BeforeUpdate", Now, Err.Number, Err.Description, True, db)
  Resume Next

End Sub

Sub CalcManyTotals()

 ' 'On Error GoTo CalcManyTotals_Error

  'Calculate totals on Pay Many Form

  Dim PaymentTotal@
  'PaymentTotal@ = IIf(IsNull(SumRecord("[PH Payment Total]", "AP Pay Many Header", "[PH Selected] = True")), 0, SumRecord("[PH Payment Total]", "AP Pay Many Header", "[PH Selected] = True"))
  'PaymentTotal@ = SumRecord("[PH Payment Total]", "AP Pay Many Header", "[PH Selected] = Yes")
  If ADOprimaryrs.RecordCount = 0 Then
    MsgBox "There is no payment to be made", vbInformation, "Payment"
    Exit Sub
  End If
  ADOprimaryrs.Update
  'adoPrimaryRS.MoveFirst
  'PaymentTotal@ = 0
  'Do While Not adoPrimaryRS.EOF
  '   If adoPrimaryRS![PH Selected] = True Then
  '       PaymentTotal@ = PaymentTotal@ + adoPrimaryRS![PH Payment Total]
  '   End If
  '   adoPrimaryRS.MoveNext
  'Loop
  'PaymentTotal@ = SumRecord("[PH Payment Total]", "AP Pay Many Header", "[PH Selected] = Yes")
  
  Dim TempRec As ADODB.Recordset
  
  PaymentTotal@ = 0
      Set TempRec = New ADODB.Recordset
        TempRec.Open "SELECT [PH Payment Total] FROM [AP Pay Many Header] WHERE [PH Selected] = Yes", db, adOpenStatic, adLockReadOnly, adCmdText
        If TempRec.RecordCount = 0 Then GoTo NoData
        TempRec.MoveFirst
        TempCurr = 0
        Do While Not TempRec.EOF
            PaymentTotal@ = PaymentTotal@ + TempRec![PH Payment Total]
            TempRec.MoveNext
        Loop
NoData:
        If TempRec.RecordCount = 0 Then PaymentTotal@ = 0
        
      TempRec.Close
      Set TempRec = Nothing
  txtFields(6) = FormatCurr(PaymentTotal@)

  txtFields(7) = txtFields(5) - txtFields(6)
  txtFields(7) = FormatCurr(txtFields(7))
  
  Exit Sub
CalcManyTotals_Error:
  Call ErrorLog("Purchase Module", "CalcManyTotals", Now, Err.Number, Err.Description, True, db)
  Resume Next
  
End Sub

Private Function CalcBalance() As Currency
  ''On Error Resume Next
  'Dim TempBalance#
  
  With DataGridDetaila
    'TempBalance# = Me![Original Amount] - Me![Discount] - Me![Write Off] - Me![Applied Amount] - Me![Amount Paid]
    CalcBalance = .Columns(3).Value - .Columns(4).Value - .Columns(5).Value - .Columns(7).Value
  End With
  'If TempBalance# < 0 Then
  '  MsgBox "Balance cannot be less than zero!", , "Error"
  '  Exit Function
  'Else
  '  DataGridDetaila.Columns(6).Value = FormatCurr(TempBalance )
  '  DataGridDetaila.col = 6
  '  SendKeys ("{ENTER}")
  'End If
End Function

Private Sub cmdDate_Click(Index As Integer)
Select Case Index
Case 1
    Menu_Calendar.WhoCallMe True, 1011
    'Menu_Calendar.Show vbModal
Case 2
    NowLoad = True
    ShowStatus True
    TempDate = txtFields(14)
    Menu_Calendar.WhoCallMe True, 1012
    'Menu_Calendar.Show vbModal
    If txtFields(14) = TempDate Then
        ShowStatus False
        Exit Sub
    End If
    strSQL = "AND [AP PO Due Date] <=#" & txtFields(14) & "# AND [AP PO Document Type] in ('Receiving','Voucher','Beginning Balance')"
    'Debug.Print strSQL
    Call LoadFirstCashAcct
    Call FillHeader
    
    dataGridSource "select * from [AP Pay Many Header]", grdDataGrid
    
    Call CalcManyTotals
    ShowStatus False
    NowLoad = False
Case 0
    NowLoad = True
    txtFields(14) = FormatDate(Now)
    ShowStatus True
    strSQL = "AND [AP PO Document Type] in ('Receiving','Voucher','Beginning Balance')"
    'Debug.Print strSQL
    Call LoadFirstCashAcct
    Call FillHeader
    
    dataGridSource "select * from [AP Pay Many Header]", grdDataGrid
   
    Call CalcManyTotals
    txtFields(14) = ""
    TempDate = txtFields(14)
    ShowStatus False
    NowLoad = False
End Select

End Sub


Private Sub DataGridDetaila_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
Dim TempCurr As Currency
   TempCurr = CalcBalance
   If TempCurr < 0 Then
      MsgBox "Balance cannot be less $0.00", vbCritical, "Error"
      Cancel = True
   ElseIf DataGridDetaila.Columns(5) = 0 Then
      MsgBox "Applied Amount cannot be $0.00", vbCritical, "Error"
      Cancel = True
   Else
      Cancel = False
      DataGridDetaila.Columns(6).Text = FormatCurr(TempCurr)
      SendKeys ("{ENTER}")
      SendKeys ("{down}")
      SendKeys ("{up}")
      
   End If

End Sub

Private Sub DataGridDetaila_Error(ByVal DataError As Integer, Response As Integer)
    If DataGridKnownError(DataError) Then
        Response = 0
    End If
End Sub

Private Sub DataGridDetaila_LostFocus()
If grdDataGrid.Row > -1 Then
    If grdDataGrid.Columns(4) <> txtFields(0) Then
        grdDataGrid.SetFocus
        grdDataGrid.Columns(4) = txtFields(0)
        SendKeys ("{ENTER}")
        SendKeys ("{down}")
        SendKeys ("{up}")
        CalcManyTotals
    End If
End If
    'adoPrimaryRS![PH Payment Total] = txtFields(0)
    'adoPrimaryRS.Update
    'dataGridSource "select * from [AP Pay Many Detail] WHERE [Vendor]='" & grdDataGrid.Columns(1).Value & "'", DataGridDetaila
End Sub

Private Sub DataGridDetaila_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

If grdDataGrid.Row > -1 Then

    Select Case DataGridDetaila.col
    Case 4
        DataGridDetaila.AllowUpdate = True
    Case 5
        DataGridDetaila.AllowUpdate = True
    Case Else
        DataGridDetaila.AllowUpdate = False
    End Select
    
    CalcAppliedUnApplied
Else
    DataGridDetaila.AllowUpdate = False
End If
End Sub

Private Sub CalcAppliedUnApplied()

Dim TempBal As Currency
Dim TempCurr As Currency
Dim TempRec As ADODB.Recordset
      
      Set TempRec = New ADODB.Recordset
        TempRec.Open "SELECT [Applied Amount],[Balance] FROM [AP Pay Many Detail] WHERE [Vendor]='" & grdDataGrid.Columns(1).Value & "'", db, adOpenStatic, adLockReadOnly, adCmdText
        TempRec.MoveFirst
        TempCurr = 0
        Do While Not TempRec.EOF
            TempCurr = TempCurr + TempRec![Applied Amount]
            TempBal = TempBal + TempRec![Balance]
            TempRec.MoveNext
        Loop
      TempRec.Close
      Set TempRec = Nothing
        txtFields(0) = FormatCurr(TempCurr)
        txtFields(2) = FormatCurr(TempBal)

End Sub

Private Sub Form_Load()
'On Error GoTo FormErr
  Set db = New ADODB.Connection
  db.CursorLocation = adUseClient
  db.Open gblADOProvider
  
'  Set adoPrimaryRS2 = New ADODB.Recordset
  NowLoad = True
  ShowStatus True
  txtFields(4).Text = FormatDate(Now)
  strSQL = "AND [AP PO Document Type] in ('Receiving','Voucher','Beginning Balance')"
  txtFields(14) = FormatDate(Now)
  Call LoadFirstCashAcct
  ShowStatus True
  Call FillHeader
  'Clear AP Pay Many Header Table
  'Build AP Pay Many Table
  
  dataGridSource "select * from [AP Pay Many Header]", grdDataGrid
  
  Call CalcManyTotals
  ShowStatus False
  NowLoad = False
  'Clear Cash Payments Work Table
  'Build Cash Payments Work Table
  
  'adoPrimaryRS2.Open "select *  from [Cash Payments Work] Order by [Reference]", db, adOpenStatic, adLockOptimistic
  
  'Dim ctrl As Control
  'For Each ctrl In Me.Controls
  '  If TypeOf ctrl Is TextBox Or TypeOf ctrl Is CheckBox Then
  '    'Set ctrl.DataSource = adoPrimaryRS2
  '  End If
  'Next
  
  GetTextColor Me
  'mbDataChanged = False
Exit Sub
FormErr:
  MsgBox Err.Description
  ShowStatus False
End Sub

Private Sub dataGridSource(SQLstatement As String, grdDtGrid As DataGrid)
  
  If grdDtGrid.Name = "grdDataGrid" Then
    Set ADOprimaryrs = New ADODB.Recordset
    TempStr = SQLstatement
    ADOprimaryrs.Open SQLstatement, db, adOpenStatic, adLockOptimistic
    Set grdDtGrid.DataSource = Nothing
    grdDtGrid.ReBind
    grdDtGrid.Refresh
    Set grdDtGrid.DataSource = ADOprimaryrs
    
    grdDataGrid.Height = 4305
    DataGridDetaila.Visible = False
    grdDtGrid.Columns(0).Button = True
  ElseIf NowLoad = False Then
    Set adoPrimaryRS2 = New ADODB.Recordset
    adoPrimaryRS2.Open SQLstatement, db, adOpenStatic, adLockOptimistic
    Set grdDtGrid.DataSource = Nothing
    'grdDtGrid.HoldFields
    grdDtGrid.Refresh
    grdDtGrid.ReBind
    Set grdDtGrid.DataSource = adoPrimaryRS2
  End If
  
End Sub

Private Sub Form_Resize()
  ''On Error Resume Next
  'This will resize the grid when the form is resized
  'grdDataGrid.Width = Me.ScaleWidth
  'grdDataGrid.Height = Me.ScaleHeight - grdDataGrid.Top - 30 - picButtons.Height - picStatBox.Height
  'lblStatus.Width = Me.Width - 1500
  'cmdNext.Left = lblStatus.Width + 700
  'cmdLast.Left = cmdNext.Left + 340
  If fMainForm.WindowState = 1 Then Exit Sub
  If Me.WindowState = 0 Then
  ElseIf Me.WindowState = 2 Then
    GoTo SkipResize
  Else
    Exit Sub
  End If
  Me.Width = 9495
  Me.Height = 7620
SkipResize:
  frPrimary.Left = (Me.ScaleWidth - frPrimary.Width) / 2
  Label1.Left = frPrimary.Left
  Label1.Width = frPrimary.Width
  If Me.Width = 9495 Then
    frPrimary.Top = 480
  Else
    frPrimary.Top = (Me.ScaleHeight - frPrimary.Height) / 2 + 230
  End If
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
'On Error GoTo FormErr
  
  'Call RedoPurchaseNumbers(db)
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
    RefreshButton ADOprimaryrs, grdDataGrid
End Sub

Private Sub cmdEdit_Click()
  ''On Error GoTo EditErr

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
  'mbDataChanged = False

End Sub

Private Sub cmdUpdate_Click()
Dim FlagStatus As Boolean
    
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
  ''On Error GoTo GoFirstError

  ADOprimaryrs.MoveFirst
  'mbDataChanged = False

  Exit Sub

GoFirstError:
  MsgBox Err.Description
End Sub

Private Sub cmdLast_Click()
  ''On Error GoTo GoLastError

  ADOprimaryrs.MoveLast
  'mbDataChanged = False

  Exit Sub

GoLastError:
  MsgBox Err.Description
End Sub

Private Sub cmdNext_Click()
  ''On Error GoTo GoNextError

  If Not ADOprimaryrs.EOF Then ADOprimaryrs.MoveNext
  If ADOprimaryrs.EOF And ADOprimaryrs.RecordCount > 0 Then
    Beep
     'moved off the end so go back
    ADOprimaryrs.MoveLast
  End If
  'show the current record
  'mbDataChanged = False

  Exit Sub
GoNextError:
  MsgBox Err.Description
End Sub

Private Sub cmdPrevious_Click()
  ''On Error GoTo GoPrevError

  If Not ADOprimaryrs.BOF Then ADOprimaryrs.MovePrevious
  If ADOprimaryrs.BOF And ADOprimaryrs.RecordCount > 0 Then
    Beep
    'moved off the end so go back
    ADOprimaryrs.MoveFirst
  End If
  'show the current record
  'mbDataChanged = False

  Exit Sub

GoPrevError:
  MsgBox Err.Description
End Sub

Private Sub SetButtons(bVal As Boolean)
  'cmdAdd.Visible = bVal
  'cmdUpdate.Visible = bVal
  'cmdCancel.Visible = Not bVal
  'cmdDelete.Visible = bVal
  'cmdClose.Visible = bVal
  'cmdRefresh.Visible = bVal
  'cmdNext.Enabled = bVal
  'cmdFirst.Enabled = bVal
  'cmdLast.Enabled = bVal
  'cmdPrevious.Enabled = bVal
End Sub

Private Sub RefreshData()

  ''On Error GoTo RefreshData_Error
  
  Dim rsWork As ADODB.Recordset
  'Dim rsCross As ADODB.Recordset
  Dim rsPurchase As ADODB.Recordset
  Dim PurchaseID$

  'Set rsWork = New ADODB.Recordset
  'rsWork.Open ("AP Pay Many Detail")
  
  'Set rsCross = New ADODB.Recordset
  'rsCross.Open "[AP Payment Invoice Cross Reference]", db, adOpenStatic, adLockOptimistic, adCmdTable
  
  Set rsPurchase = New ADODB.Recordset
  rsPurchase.Open "SELECT [AP PO Ext Document No],[AP PO Document No] FROM [AP Purchase]", db, adOpenKeyset, adLockOptimistic, adCmdText

  'rsPurchase.Index = "Ext Document #"

  'Dim rsPayment As ADODB.Recordset
  'Set rsPayment = New ADODB.Recordset
  'rsPayment.Open "SELECT * FROM [AP Payment Header] WHERE [AP PAY Transaction Date]=" & txtFields(4), db, adOpenStatic, adLockOptimistic, adCmdText

  Dim ID&
  Dim TotalApplied#

  Dim rsPayMany As ADODB.Recordset
    'If optSortOrder.Value = True Then
        'Order by Vendor Name
        Set rsPayMany = New ADODB.Recordset
        rsPayMany.Open "SELECT [PH Check Number],[PH Vendor ID],[PH Payment Total],[PH Vendor ID] FROM [qryPayManyHeader] where [PH Payment Total] > 0", db, adOpenStatic, adLockOptimistic, adCmdText
    'Else
        'Order by Zip Code
        'Set rsPayMany = New ADODB.Recordset
        'rsPayMany.Open "SELECT [PH Check Number],[PH Vendor ID],[PH Payment Total],[PH Vendor ID] FROM [qryPayManyHeader] where [PH Payment Total] > 0 ORDER BY [AP VEN Postal] Asc", db, adOpenStatic, adLockOptimistic, adCmdText
    'End If
  '
  rsPayMany.MoveFirst
  Do While Not rsPayMany.EOF
    'Create a cash payment record from this
    'rsPayMany.Edit
      rsPayMany("PH Check Number") = NextCheck$
    rsPayMany.Update
    'guna insertSQLcommand
    'rsPayMany.Update
      
      SQLstatement = "INSERT INTO [AP Payment Header]"
      SQLstatement = SQLstatement & " ([AP PAY Type],[AP PAY Check No]," & _
      "[AP PAY Vendor No],[AP PAY Transaction Date],[AP PAY Amount],[AP PAY UnApplied Amount]," & _
      "[AP PAY Bank Account],[AP PAY Void],[AP PAY Notes],[AP PAY Credit Amount]," & _
      "[AP PAY Class],[AP PAY Cleared],[AP PAY Posted YN],[AP PAY Recurring YN],[AP PAY Status])"
      SQLstatement = SQLstatement & " VALUES ('Payment','" & NextCheck$ & "','" & _
      rsPayMany("PH Vendor ID") & "" & "',#" & txtFields(4).Text & "#," & rsPayMany("PH Payment Total") & ",0,'" & _
      txtFields(1).Text & "',False,'Paid through pay many.',0," & _
      "0,False,False,False,'Open')"
      db.Execute SQLstatement
    
    'rsPayment.AddNew
    '  rsPayment("AP PAY Type") = "Payment"
    '  rsPayment("AP PAY Check No") = NextCheck$
    '  rsPayment("AP PAY Vendor No") = rsPayMany("PH Vendor ID") & ""
    '  rsPayment("AP PAY Transaction Date") = txtFields(4)
    '  rsPayment("AP PAY Amount") = rsPayMany("PH Payment Total")
    '  rsPayment("AP PAY UnApplied Amount") = 0
    '  rsPayment("AP PAY Bank Account") = txtFields(1) & ""
    '  rsPayment("AP PAY Status") = "" '-----
    '  rsPayment("AP PAY Void") = False
    '  rsPayment("AP PAY Notes") = "Paid through pay many."
    '  rsPayment("AP PAY Credit Amount") = 0
    '  rsPayment("AP PAY Class") = 0
    '  rsPayment("AP PAY Cleared") = False
    '  rsPayment("AP PAY Posted YN") = False
    '  rsPayment("AP PAY Recurring YN") = False
    '  rsPayment("AP PAY Status") = "Open"
    'rsPayment.Update
    
    Dim TemprsPayment As ADODB.Recordset
    Set TemprsPayment = New ADODB.Recordset
    TemprsPayment.Open "SELECT [AP PAY ID] FROM [AP Payment Header] WHERE [AP PAY Check No]='" & NextCheck$ & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    ID& = TemprsPayment("AP PAY ID")
    TemprsPayment.Close
    Set TemprsPayment = Nothing
    
    'Save detail information
    ''On Error Resume Next
    TotalApplied# = 0
    Set rsWork = New ADODB.Recordset
    rsWork.Open "SELECT * FROM [AP Pay Many Detail] where [Vendor] = '" & rsPayMany("PH Vendor ID") & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    If rsWork.RecordCount > 0 Then
      Do While Not rsWork.EOF
        PurchaseID$ = rsWork("Reference")
        rsPurchase.MoveFirst
        'rsPurchase.Seek "=", PurchaseID$
        rsPurchase.Find "[AP PO Ext Document No]='" & PurchaseID$ & "'"
        If rsPurchase.EOF Then
          'Should not happen
        Else
          If rsWork("Applied Amount") > 0 Then  'rsWork("Discount") > 0 Or rsWork("Write Off") > 0 Or rsWork("Applied Amount") > 0 Then
            'Write a cross reference record
            SQLstatement = "INSERT INTO [AP Payment Invoice Cross Reference]"
            SQLstatement = SQLstatement & " ([AP CROSS Payment ID],[AP CROSS Payed ID]," & _
            "[AP CROSS Discount Taken],[AP CROSS Write Off Amount],[AP CROSS Applied Amount],[AP CROSS Cleared])"
            SQLstatement = SQLstatement & " VALUES (" & ID& & "," & rsPurchase("AP PO Document No") & "," & _
            rsWork("Discount") & "," & rsWork("Write Off") & "," & rsWork("Applied Amount") & ",False)"
            db.Execute SQLstatement
            'rsCross.AddNew
            '  rsCross("AP CROSS Payment ID") = ID&
            '  rsCross("AP CROSS Payed ID") = rsPurchase("AP PO Document No")
            '  rsCross("AP CROSS Discount Taken") = rsWork("Discount")
            '  rsCross("AP CROSS Write Off Amount") = rsWork("Write Off")
            '  rsCross("AP CROSS Applied Amount") = rsWork("Applied Amount")
            '  TotalApplied# = TotalApplied# + rsWork("Applied Amount")
            '  rsCross("AP CROSS Cleared") = False
            'rsCross.Update
          Else
          End If
        End If
        rsWork.MoveNext
      Loop
    End If
    NextCheck$ = Trim(CStr(Val(NextCheck$) + 1))
    rsPayMany.MoveNext
  Loop

  Exit Sub
RefreshData_Error:
  Call ErrorLog("Pay Many", "RefreshData", Now, Err.Number, Err.Description, True, db)
  Resume Next

End Sub

Private Sub PostData()

  ''On Error GoTo PostData_Error

  Dim rsWork As ADODB.Recordset
  'Dim rsCross As ADODB.Recordset
  Dim rsPurchase As ADODB.Recordset
  Dim PurchaseID$
  
  
  'xxx 1/3/97 7.2b
  Dim sql$
  sql$ = "SELECT DISTINCTROW [AP Pay Many Header].[PH Payment Total], [AP Pay Many Header].[PH Selected], [AP Pay Many Detail].*"
  sql$ = sql$ & " FROM [AP Pay Many Header] INNER JOIN [AP Pay Many Detail] ON [AP Pay Many Header].[PH Vendor ID] = [AP Pay Many Detail].Vendor"
  sql$ = sql$ & " WHERE ((([AP Pay Many Header].[PH Payment Total])>0) AND (([AP Pay Many Header].[PH Selected])=True))"

  Set rsWork = New ADODB.Recordset
  rsWork.Open sql$, db, adOpenKeyset, adLockOptimistic, adCmdText

'  Set rsWork = db2.OpenRecordset("AP Pay Many Detail")
  
  'Set rsCross = New ADODB.Recordset
  'rsCross.Open "[AP Payment Invoice Cross Reference]"
  Set rsPurchase = New ADODB.Recordset
  rsPurchase.Open "SELECT [AP PO Ext Document No],[AP PO Amount Paid],[AP PO Total Amount], " & _
  "[AP PO Balance Due] FROM [AP Purchase]", db, adOpenKeyset, adLockOptimistic, adCmdText

  'rsPurchase.Index = "Ext Document #"

  'Save detail information
  ''On Error Resume Next
  If rsWork.RecordCount > 0 Then
    Do While Not rsWork.EOF
      PurchaseID$ = rsWork("Reference")
      'rsPurchase.Seek "=", PurchaseID$
      rsPurchase.MoveFirst
      rsPurchase.Find "[AP PO Ext Document No]='" & PurchaseID$ & "'"
      If rsPurchase.EOF Then
        'Should not happen
      Else
        'Update Purchase Record
        If rsWork("Applied Amount") > 0 Then
          'rsPurchase.Edit
            rsPurchase("AP PO Amount Paid") = rsPurchase("AP PO Total Amount") - rsWork("Balance")
            rsPurchase("AP PO Balance Due") = rsWork("Balance")
          rsPurchase.Update
        End If
      End If
      rsWork.MoveNext
    Loop
  End If
  
  rsWork.Close
  Set rsWork = Nothing
  
  rsPurchase.Close
  Set rsPurchase = Nothing
  
  Exit Sub
PostData_Error:
  Call ErrorLog("Pay Many", "PostData", Now, Err.Number, Err.Description, True, db)
  Resume Next
  
End Sub

Public Sub LoadFirstCashAcct()
'Dim rsAP As ADODB.Recordset

  'Set rsAP = New ADODB.Recordset
  'rsAP.Open "SELECT * FORM [Cash Payments Work] WHERE [Reference]='" & "000" & "'", db, adOpenStatic, adLockOptimistic, adCmdText
  'rsAP.AddNew
  '  rsAP("Reference") = "000"
  'rsAP.Update
  'rsAP.Close
  db.Execute "INSERT INTO [Cash Payments Work] ([Reference]) VALUES ('000')"
  
  'Me.RecordLocks = 0
  'DoCmd.GoToRecord A_FORM, "Pay Many", acFirst
  'Me.RecordLocks = 1

  Dim rsGL As ADODB.Recordset
  Set rsGL = New ADODB.Recordset
  rsGL.Open "SELECT [GL COA Account No],[GL COA Account Balance] FROM [GL Chart Of Accounts] where [GL COA Asset Type] = 'Cash'", db, adOpenStatic, adLockOptimistic, adCmdText

  ''On Error Resume Next
  'Err = 0
  If rsGL.RecordCount = 0 Then
    MsgBox "There is no default cash account to post to.", vbInformation, "Information"
  Else
    rsGL.MoveFirst
    txtFields(1) = rsGL("GL COA Account No")
    'txtFields(5) = rsGL("GL COA Account Balance")
    txtFields(5) = FormatCurr(rsGL("GL COA Account Balance"))
    'Dim rsBank As ADODB.Recordset
    'Set rsBank = New ADODB.Recordset
    'rsBank.Open "SELECT * FROM [Bank Accounts] WHERE [BANK ACCT ID]='" & txtFields(1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    'rsBank.Index = "PrimaryKey"
    'rsBank.Seek "=", Me![Bank Acct]
    'txtFields(8) = rsBank![BANK ACCT Next Check No]
    
    'txtFields(5) = AcctBalance("balance", txtFields(1).Text)
    
    'get a valid checkNumber
    Dim GetCheckNo As String
    GetCheckNo = CheckNumberCHQ("read", db, txtFields(1).Text)
    If GetCheckNo <> "" Then
      txtFields(8).Text = GetCheckNo
    Else
      txtFields(8).Text = ""
    End If
  End If

End Sub
Private Sub FillHeader()

  ''On Error GoTo FillHeader_Error

  'Fill Pay Many Header with information
  db.Execute "DELETE * FROM [AP Pay Many Header]"
  db.Execute "DELETE * FROM [AP Pay Many Detail]"

  Dim rsPurchase As ADODB.Recordset
  Dim rsWork As ADODB.Recordset
  'Dim rsHeader As ADODB.Recordset
  Dim rsCross As ADODB.Recordset
  Dim rsVendor As ADODB.Recordset
  Dim TotalBalance@

  'Set rsHeader = New ADODB.Recordset
  'rsHeader.Open "SELECT * FROM [AP Pay Many Header]", db, adOpenKeyset, adLockOptimistic, adCmdText
  
  Set rsWork = New ADODB.Recordset
  rsWork.Open "SELECT * FROM [AP Pay Many Detail]", db, adOpenKeyset, adLockOptimistic, adCmdText

  Dim PurchaseID&
  Dim VendorID$
  Dim DueDate As String

  Set rsVendor = New ADODB.Recordset
  rsVendor.Open "SELECT [AP VEN ID],[AP VEN Name] FROM [AP Vendor]", db, adOpenKeyset, adLockOptimistic, adCmdText
  rsVendor.MoveFirst
  
  Dim SQLstatement As String
  Do While Not rsVendor.EOF
    VendorID$ = rsVendor("AP VEN ID")
    TotalBalance@ = 0
    
    'Load all invoices for this vendor w/ a balance
    Set rsPurchase = New ADODB.Recordset
    SQLstatement = "SELECT [AP PO Document No],[AP PO Date],[AP PO Ext Document No],[AP PO Vendor Invoice No]," & _
    "[AP PO Document Type],[AP PO Due Date],[AP PO Balance Due],[AP PO Amount Paid], " & _
    "[AP PO Total Amount] FROM [AP Purchase] WHERE [AP PO Posted YN] = True and [AP PO Balance Due] > 0 and [AP PO Vendor ID] ='" & VendorID$ & "' " & strSQL
    'Debug.Print SQLStatement
    rsPurchase.Open SQLstatement, db, adOpenKeyset, adLockOptimistic, adCmdText
    ''On Error Resume Next
    If rsPurchase.RecordCount > 0 Then
      rsPurchase.MoveFirst
      'Load these records into work table
      Do While Not rsPurchase.EOF
      rsWork.AddNew
          rsWork("Vendor") = VendorID$
          If IsNull(rsPurchase("AP PO Vendor Invoice No")) Then
            rsWork("AP PO Vendor Invoice No") = "Not Known"
          Else
            rsWork("AP PO Vendor Invoice No") = rsPurchase("AP PO Vendor Invoice No") & ""
          End If
          rsWork("AP PAY Type") = rsPurchase("AP PO Document Type") & ""
          rsWork("AP PO Due Date") = rsPurchase("AP PO Due Date")
          DueDate = rsPurchase("AP PO Due Date")
          rsWork("Reference") = rsPurchase("AP PO Ext Document No") & ""
          rsWork("Date") = rsPurchase("AP PO Date")
          rsWork("Original Amount") = rsPurchase("AP PO Total Amount")
          rsWork("Amount Paid") = rsPurchase("AP PO Amount Paid")
          If rsWork("Amount Paid") > 0 Then
            rsWork("Discount") = 0
          Else
            rsWork("Discount") = Round(GetAPInvoiceDiscount(CLng(rsPurchase("AP PO Document No")), txtFields(4)))
          End If
          rsWork("Write Off") = 0
          rsWork("Applied Amount") = rsPurchase("AP PO Balance Due") - rsWork("Discount")
          rsWork("Balance") = 0 'rsPurchase("AP PO Balance Due") - rsWork("Discount")
          TotalBalance@ = TotalBalance@ + rsWork("Applied Amount")
      rsWork.Update
      rsPurchase.MoveNext
      Loop
    End If
    'Create header record
    If TotalBalance@ > 0 Then
    
      SQLstatement = "INSERT INTO [AP Pay Many Header]"
      SQLstatement = SQLstatement & " ([PH Vendor ID],[PH Vendor Name],[PH Total Owed],[PH Payment Total],[PH Selected],[PH Due Date])"
      SQLstatement = SQLstatement & " VALUES ('" & VendorID$ & "','" & rsVendor("AP VEN Name") & "" & "'," & TotalBalance@ & "," & TotalBalance@ & ",True,#" & DueDate & "#)"
      db.Execute SQLstatement
      
      'rsHeader.AddNew
      '  rsHeader("PH Vendor ID") = VendorID$
      '  rsHeader("PH Vendor Name") = rsVendor("AP VEN Name") & ""
      '  rsHeader("PH Total Owed") = TotalBalance@
      '  rsHeader("PH Payment Total") = TotalBalance@
      '  rsHeader("PH Selected") = True
      '  rsHeader("PH Due Date") = DueDate
        
      'rsHeader.Update
    End If
    rsVendor.MoveNext
    
    rsPurchase.Close
    Set rsPurchase = Nothing
  Loop

  'Refresh the tables
  rsWork.Close
  Set rsWork = Nothing
  'rsHeader.Close
  'Set rsHeader = Nothing
  
  Exit Sub
FillHeader_Error:
  Call ErrorLog("Pay Many", "FillHeader", Now, Err.Number, Err.Description, True, db)
  Resume Next
  
End Sub

Private Sub cmdPost_Click()

  ''On Error GoTo cmdPost_Click_Error

  'Get the next check number and print each check
  'Write records to AP Payment Header and Detail
  'Post these records
  NowLoad = True
  
  If Trim(txtFields(8)) = "" Then
    MsgBox "Transaction cancel! Please Enter Check Number", vbInformation, "Error"
    Exit Sub
  End If
  If CCur(txtFields(7)) < 0 Then
    MsgBox "End Balance cannot be less than $0.00", vbInformation, "Information"
    Exit Sub
  End If
  
  '---------------------- use adoprimaryrs.recordcount
  Dim rsPayMany As ADODB.Recordset
  Set rsPayMany = New ADODB.Recordset
  rsPayMany.Open "SELECT [PH Payment Total] FROM [AP Pay Many Header] where [PH Payment Total] > 0 and [PH Selected] = TRUE", db, adOpenKeyset, adLockReadOnly, adCmdText
  If rsPayMany.RecordCount = 0 Then
    MsgBox "There are no checks to print!", , "Error"
    Exit Sub
  End If
  rsPayMany.Close
  Set rsPayMany = Nothing
  '---------------------------------------------------
  
  'Get the next check no
  
  Dim rsBank As ADODB.Recordset
  'Set rsBank = New ADODB.Recordset
  'rsBank.Open "SELECT [BANK ACCT Next Check No] FROM [Bank Accounts] WHERE [BANK ACCT ID]='" & txtFields(1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
  'rsBank.Index = "PrimaryKey"
  'rsBank.Seek "=", Me![Bank Acct]
  'If rsBank.RecordCount = 0 Then
  '  MsgBox "Bank account is not valid!", , "Error"
  '  Exit Sub
  'End If
  
  '------------------ start---------------------------------
  
  Dim VendorID$
  Dim BankID$
  Dim CheckNo$
  Dim NumChecks%
  Dim rsHeader As ADODB.Recordset
  Dim ThisCheck&


GetCheckRange:

  NextCheck$ = txtFields(8)
  'NextCheck$ = rsBank("BANK ACCT Next Check No")
  'NextCheck$ = InputBox("Enter next check number.", "Check", NextCheck$)
  'If NextCheck$ = "" Or NextCheck$ = txtFields(8) Then Exit Sub
  If Not IsNumeric(NextCheck$) Then
    MsgBox "Please enter a valid check number!", , "Error"
    GoTo GetCheckRange
  End If

  ShowStatus True

  Set rsHeader = New ADODB.Recordset
  rsHeader.Open "SELECT [AP PAY Check No],[AP PAY Void] FROM [AP Payment Header]WHERE [AP PAY Bank Account]='" & txtFields(1) & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
  'rsHeader.Index = "BankKey"
  'Make sure all checks in the desired range are not already used
  'If optSortOrder = 1 Then
    'Order by Vendor Name
    Set rsPayMany = New ADODB.Recordset
    'rsPayMany.Open "SELECT * FROM [qryPayManyHeader] where [PH Payment Total] > 0 ORDER BY [AP VEN Name] Asc", db, adOpenStatic, adLockOptimistic, adCmdText
    rsPayMany.Open "SELECT [PH Vendor ID],[PH Check Number] FROM [qryPayManyHeader] where [PH Payment Total] > 0", db, adOpenStatic, adLockOptimistic, adCmdText
    'MsgBox rsPayMany.RecordCount
  'Else
    'Order by Zip Code
    'Set rsPayMany = New ADODB.Recordset
    'rsPayMany.Open "SELECT * FROM [qryPayManyHeader] where [PH Payment Total] > 0 ORDER BY [AP VEN Postal] Asc", db, adOpenStatic, adLockOptimistic, adCmdText
  'End If
'again:
  rsPayMany.MoveFirst
  ThisCheck& = Val(NextCheck$)
  
  Do While Not rsPayMany.EOF
    VendorID$ = rsPayMany("PH Vendor ID")
'GetNewCheck:
    CheckNo$ = Trim(CStr(ThisCheck&))
    BankID$ = CStr(txtFields(1))
    'GoTo again
    ThisCheck& = CheckNumberCHQ("read", db, BankID$, CheckNo$)
    rsPayMany![PH Check Number] = ThisCheck&
    'rsHeader.MoveFirst
    'rsHeader.Find "[AP PAY Check No]='" & CheckNo$ & "'"
    '    If rsHeader.EOF Then
          'This is a good thing
    '    Else
    '      ShowStatus False
    '      ThisCheck& = InputBox("A check " & ThisCheck& & " has already been used" & vbCr & "Enter New check number.", "Check", ThisCheck&)
          'MsgBox "A check in the desired range has already been used!", , "Error"
    '      GoTo GetNewCheck
    '    End If
    rsPayMany.Update
    ThisCheck& = ThisCheck& + 1
    rsPayMany.MoveNext
  Loop
  'rsPayMany.Close

  Dim Response%
  Response% = MsgBox("Make sure checks are in the printer and press OK.", vbOKCancel, "Information")
  If Response% = vbCancel Then
    ShowStatus False
    Exit Sub
  End If

  Call RefreshData
  
  Dim Success%

  'Now print the checks
  'If optSortOrder = 1 Then
    'Order by Vendor Name
  '  Set rsPayMany = New ADODB.Recordset
  '  rsPayMany.Open "SELECT * FROM [qryPayManyHeader] where [PH Payment Total] > 0 ORDER BY [AP VEN Name] Asc", db, adOpenStatic, adLockOptimistic, adCmdText
    'db.OpenRecordset ("SELECT * FROM [qryPayManyHeader] where [PH Payment Total] > 0 ORDER BY [AP VEN Name] Asc")
  'Else
    'Order by Zip Code
  '  Set rsPayMany = New ADODB.Recordset
    'db.OpenRecordset("SELECT * FROM [qryPayManyHeader] where [PH Payment Total] > 0 ORDER BY [AP VEN Postal] Asc")
  '  rsPayMany.Open "SELECT * FROM [qryPayManyHeader] where [PH Payment Total] > 0 ORDER BY [AP VEN Postal] Asc", db, adOpenStatic, adLockOptimistic, adCmdText
  'End If
  'rsPayMany.Requery
  rsPayMany.MoveFirst
  Do While Not rsPayMany.EOF
    VendorID$ = rsPayMany("PH Vendor ID")
    CheckNo$ = rsPayMany("PH Check Number")
    BankID$ = CStr(txtFields(1))
    Success% = PrintCheck(VendorID$, CheckNo$, BankID$, txtFields(4), db)
    rsPayMany.MoveNext
  Loop

  Dim PayID&
'
  'Dim rsHeader As ADODB.Recordset
  'Set rsHeader = New ADODB.Recordset
  'rsHeader.Open "SELECT * FROM [AP Payment Header]WHERE [AP PAY Bank Account]='" & txtFields(1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
  'rsHeader.Index = "BankKey"

  Response% = MsgBox("Did the checks print correctly?", vbYesNo)
  If Response% = vbNo Then
    'Mark printed checks as void
    Dim Buffer%
    rsPayMany.MoveFirst
    Do While Not rsPayMany.EOF
      VendorID$ = rsPayMany("PH Vendor ID")
      CheckNo$ = rsPayMany("PH Check Number")
      BankID$ = CStr(txtFields(1))
      'rsHeader.Seek "=", CheckNo$, BankID$
      rsHeader.Find "[AP PAY Check No]='" & CheckNo$ & "'"
      'rsHeader.Edit
      rsHeader("AP PAY Void") = True
      rsHeader.Update
      rsPayMany.MoveNext
    Loop
    
    'Reset next check number
    'rsBank.Edit
    'rsBank("BANK ACCT Next Check No") = Val(rsBank("BANK ACCT Next Check No")) + CountRecord("[PH Vendor ID]", "[AP Pay Many Header]", "[PH Payment Total] > 0 and [PH Selected] = TRUE")
    'rsBank.Update
    txtFields(8) = rsPayMany("PH Check Number") + 1 ' Val(rsBank("BANK ACCT Next Check No")) + CountRecord("[PH Vendor ID]", "[AP Pay Many Header]", "[PH Payment Total] > 0 and [PH Selected] = TRUE") - 1
    MsgBox "Your Checks have been Voided"
    Exit Sub
  Else

  End If

  '---------------------second Phase------------------------------
  
  ShowStatus True
  
  rsHeader.Close
  Set rsHeader = Nothing
  
  rsPayMany.Close
  
  'Post the data
  'Start a transaction
  Set rsPayMany = New ADODB.Recordset
  rsPayMany.Open "SELECT [PH Vendor ID],[PH Check Number] FROM [AP Pay Many Header] where [PH Payment Total] > 0 and [PH Selected] = TRUE", db, adOpenStatic, adLockOptimistic, adCmdText
  rsPayMany.MoveFirst
  Do While Not rsPayMany.EOF
    VendorID$ = rsPayMany("PH Vendor ID")
    CheckNo$ = rsPayMany("PH Check Number")
    BankID$ = txtFields(1)
    
    db.BeginTrans
    Success% = PostPayments(VendorID$, CheckNo$, BankID$, db)
    If Success% = False Then
        MsgBox "An error occured while posting the payment", vbCritical, "Information"
        db.RollbackTrans
    End If
    db.CommitTrans
    
    rsPayMany.MoveNext
  Loop
  
  'Set rsPayMany = New ADODB.Recordset
  'rsPayMany.Open "SELECT * FROM [AP Pay Many Header] where [PH Payment Total] > 0 and [PH Selected] = TRUE", db, adOpenStatic, adLockOptimistic, adCmdText
  Dim rsPayment As ADODB.Recordset
  Set rsPayment = New ADODB.Recordset
  rsPayment.Open "SELECT [AP PAY ID],[AP PAY Posted YN],[AP PAY Check No] " & _
  "FROM [AP Payment Header]WHERE [AP PAY Bank Account]='" & txtFields(1) & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
  'rsPayment.Index = "BankKey"
  rsPayMany.MoveFirst
  Do While Not rsPayMany.EOF
    VendorID$ = rsPayMany("PH Vendor ID")
    CheckNo$ = rsPayMany("PH Check Number")
    BankID$ = txtFields(1)
    rsPayment.MoveFirst
    rsPayment.Find "[AP PAY Check No]='" & CheckNo$ & "'"
    'rsPayment.Seek "=", CheckNo$, BankID$
    'rsPayment.Edit
      rsPayment("AP PAY Posted YN") = True
    rsPayment.Update
    rsPayMany.MoveNext
  Loop
  
  rsPayMany.Close
  Set rsPayMany = Nothing
  
  Call PostData
  
  'rsBank.Edit
   ' rsBank("BANK ACCT Next Check No") = Trim(CStr(Val(CheckNo$) + 1))
  'rsBank.Update
  
  'rsBank.Close
  'Set rsBank = Nothing
  
  txtFields(8) = Trim(CStr(Val(CheckNo$) + 1))
  ShowStatus True
    
  MsgBox "Transaction Posted!"
  
  'Dim rsGL As ADODB.Recordset
  'Set rsGL = New ADODB.Recordset
  'rsGL.Open "SELECT * FROM [GL Chart Of Accounts] where [GL COA Account No] = '" & txtFields(1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText

  'rsGL.MoveFirst
  'Me![Bank Acct] = rsGL("GL COA Account No")
  'txtFields(5) = FormatCurr(rsGL("GL COA Account Balance") )
  
  txtFields(5) = AcctBalance("balance", txtFields(1).Text, db)
  ShowStatus True
  Call FillHeader

  Call CalcManyTotals
  
  dataGridSource "select * from [AP Pay Many Header]", grdDataGrid

  ShowStatus False
  NowLoad = False
  
Exit Sub
cmdPost_Click_Error:
  Call ErrorLog("Pay Many", "cmdPost_Click", Now, Err.Number, Err.Description, True, db)
  Resume Next

End Sub
Private Sub grdDataGrid_ButtonClick(ByVal ColIndex As Integer)
On Error GoTo Error_ButtClick
   Select Case ColIndex
   Case 0
      If grdDataGrid.Columns(0).Text = "No" Then
         grdDataGrid.Columns(0).Text = "Yes"
      Else
         grdDataGrid.Columns(0).Text = "No"
      End If
         SendKeys ("{ENTER}")
         'adoPrimaryRS.Update
         SendKeys ("{down}")
         SendKeys ("{up}")
         
      CalcManyTotals
   End Select
Exit Sub
Error_ButtClick:
    MsgBox "Please click the Table box before clicking the button"
End Sub

Private Sub grdDataGrid_HeadClick(ByVal ColIndex As Integer)
If ADOprimaryrs.RecordCount = 0 Then Exit Sub
    ADOprimaryrs.Close
    Set ADOprimaryrs = Nothing
    Set ADOprimaryrs = New ADODB.Recordset
    'MsgBox grdDatagrid3.Columns(ColIndex).DataField
    ADOprimaryrs.Open TempStr & " ORDER BY [" & grdDataGrid.Columns(ColIndex).DataField & "]", db, adOpenKeyset, adLockOptimistic, adCmdText
    Set grdDataGrid.DataSource = ADOprimaryrs
End Sub


Private Sub grdDataGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   If grdDataGrid.Row > -1 And NowLoad = False Then
      grdDataGrid.Height = 2745
      DataGridDetaila.Visible = True
      
      dataGridSource "select * from [AP Pay Many Detail] WHERE [Vendor]='" & grdDataGrid.Columns(1).Value & "'", DataGridDetaila
      txtFields(0) = FormatCurr(grdDataGrid.Columns(4))
   End If
End Sub





Private Sub txtfields_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
Case 8
    keyResponse = CtrlValidate(KeyAscii, "0123456789")
    If keyResponse = True Then
    Else
       KeyAscii = 0
    End If
End Select
End Sub

Private Sub txtFields_LostFocus(Index As Integer)
Select Case Index
Case 8
    If IsNumeric(txtFields(8).Text) Then
        ShowStatus True
        If CheckCheckNumber(txtFields(1).Text, txtFields(8).Text, db, True) = "Found" Then
            Response% = MsgBox("Check Number is already used. Would you like to open Check Management?", vbYesNo, "Information")
            If Response% = vbYes Then
                frm_Check_Management.OpenPosted txtFields(8).Text
                End If
            txtFields(5).Text = ""
            ShowStatus False
            Exit Sub
        End If
    Else
        MsgBox "Only numeric character accepted as a check number", vbInformation, "Information"
    End If
End Select
End Sub


