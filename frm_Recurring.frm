VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_Recurring 
   Caption         =   "Recurring"
   ClientHeight    =   5340
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12975
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5340
   ScaleWidth      =   12975
   Begin VB.Frame frPrimary 
      Height          =   4815
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   12975
      Begin VB.Frame frsales 
         Height          =   2295
         Left            =   240
         TabIndex        =   13
         Top             =   2280
         Width           =   2655
         Begin VB.CommandButton cmdCreate 
            Caption         =   "Create GL Entry"
            Height          =   375
            Left            =   120
            TabIndex        =   20
            Top             =   1800
            Width           =   2415
         End
         Begin VB.CommandButton cmdDate 
            Height          =   285
            Index           =   1
            Left            =   2160
            Picture         =   "frm_Recurring.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   1320
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
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   1320
            Width           =   1215
         End
         Begin VB.CommandButton cmdDate 
            Height          =   285
            Index           =   0
            Left            =   2160
            Picture         =   "frm_Recurring.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   840
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
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Please select a date to create an invoices with next execution date:"
            Height          =   495
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   2415
         End
         Begin VB.Label lblLabels 
            Caption         =   "End Date:"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   19
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label lblLabels 
            Caption         =   "Start Date:"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   18
            Top             =   840
            Width           =   855
         End
      End
      Begin VB.ComboBox cbPurchase 
         DataField       =   "AR SALE Recur Type"
         Height          =   315
         Index           =   15
         ItemData        =   "frm_Recurring.frx":0614
         Left            =   3120
         List            =   "frm_Recurring.frx":0616
         TabIndex        =   10
         Text            =   "cbPurchase"
         Top             =   120
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Frame Frame1 
         Height          =   4575
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   2895
         Begin VB.Frame Frame3 
            Height          =   1335
            Left            =   120
            TabIndex        =   6
            Top             =   120
            Width           =   2655
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
               Left            =   120
               TabIndex        =   8
               Top             =   720
               Width           =   1335
            End
            Begin VB.CommandButton cmdSearch 
               Caption         =   "&Search"
               Height          =   855
               Left            =   1560
               Picture         =   "frm_Recurring.frx":0618
               Style           =   1  'Graphical
               TabIndex        =   7
               Top             =   240
               Width           =   855
            End
            Begin VB.Label lblfields 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Check No"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   9
               Top             =   480
               Width           =   1335
            End
         End
         Begin VB.Label lblfields 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Recurring Type"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   15
            Left            =   1440
            TabIndex        =   11
            Top             =   1200
            Visible         =   0   'False
            Width           =   1335
         End
      End
      Begin MSDataGridLib.DataGrid grdDataGrid2 
         Height          =   855
         Left            =   3120
         TabIndex        =   1
         Top             =   1800
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   1508
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
         Caption         =   "Recurring Sales"
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "AR SALE Recur Type"
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
         BeginProperty Column01 
            DataField       =   "AR SALE Next Recur"
            Caption         =   "Next Date"
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
            DataField       =   "AR SALE Ext Document #"
            Caption         =   "Document No."
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
            Caption         =   "Document Type"
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
         BeginProperty Column05 
            DataField       =   "AR SALE Total"
            Caption         =   "Sales Amount"
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
            MarqueeStyle    =   4
            BeginProperty Column00 
               Button          =   -1  'True
               ColumnWidth     =   1244.976
            EndProperty
            BeginProperty Column01 
               Button          =   -1  'True
               ColumnWidth     =   1244.976
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1755.213
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1530.142
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1665.071
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1739.906
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid grdDataGrid3 
         Height          =   855
         Left            =   3120
         TabIndex        =   2
         Top             =   2760
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   1508
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
         Caption         =   "Recurring Purchases"
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "AP PO Recurring YN"
            Caption         =   "Recurring"
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
            DataField       =   "AP PO Ext Document No"
            Caption         =   "Document No"
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
            DataField       =   "AP PO Document Type"
            Caption         =   "Document Type"
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
            DataField       =   "AP PO Date"
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
         BeginProperty Column04 
            DataField       =   "AP PO Vendor ID"
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
         BeginProperty Column05 
            DataField       =   "AP PO Total Amount"
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
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   4
            BeginProperty Column00 
               Button          =   -1  'True
               ColumnWidth     =   884.976
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1814.74
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1244.976
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1725.165
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1739.906
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid grdDataGrid4 
         Bindings        =   "frm_Recurring.frx":0922
         Height          =   855
         Left            =   3120
         TabIndex        =   3
         Top             =   3720
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   1508
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
         Caption         =   "Recurring Payments"
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "AP PAY Recurring YN"
            Caption         =   "Recurring"
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
         BeginProperty Column02 
            DataField       =   "AP PAY Type"
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
         BeginProperty Column03 
            DataField       =   "AP PAY Transaction Date"
            Caption         =   "Transaction Date"
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
         BeginProperty Column04 
            DataField       =   "AP PAY Vendor No"
            Caption         =   "Vendor No"
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
            DataField       =   "AP PAY Amount"
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
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   4
            BeginProperty Column00 
               Button          =   -1  'True
               ColumnWidth     =   884.976
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1604.976
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1454.74
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1739.906
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid grdDataGrid1 
         Height          =   4455
         Left            =   3120
         TabIndex        =   4
         Top             =   240
         Width           =   9735
         _ExtentX        =   17171
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
         Caption         =   "Recurring General Ledger"
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "GL TRANS Recurring YN"
            Caption         =   "Recurring"
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
            DataField       =   "GL TRANS Document #"
            Caption         =   "Document No."
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
            DataField       =   "GL TRANS Type"
            Caption         =   "Document Type"
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
            DataField       =   "GL TRANS Date"
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
         BeginProperty Column04 
            DataField       =   "GL TRANS Reference"
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
         BeginProperty Column05 
            DataField       =   "GL TRANS Amount"
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
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   4
            BeginProperty Column00 
               Button          =   -1  'True
               ColumnWidth     =   884.976
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1814.74
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1739.906
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
         EndProperty
      End
   End
   Begin VB.Label lblTop 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Project Types"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frm_Recurring"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ADOprimaryrs As ADODB.Recordset
Dim db As ADODB.Connection
Dim TempStr As String
Dim WhichField As String

Public Sub RequestType(MainRequest As Integer)

Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open gblADOProvider

If ADOprimaryrs Is Nothing Then
Else
    ADOprimaryrs.Close
    Set ADOprimaryrs = Nothing
End If

Set ADOprimaryrs = New ADODB.Recordset

cmdDate(0).Enabled = False
cmdDate(1).Enabled = False
'cmdCreate.Enabled = False
lblLabels(3).Enabled = False
lblLabels(4).Enabled = False
txtfields(0).Enabled = False
txtfields(1).Enabled = False
Label1.Enabled = False

Select Case MainRequest
Case 1
    TempStr = "SELECT [GL TRANS Number],[GL TRANS Recurring YN],[GL TRANS Document #],[GL TRANS Type],[GL TRANS Date],[GL TRANS Reference],[GL TRANS Amount] FROM [GL Transaction] WHERE [GL TRANS Recurring YN] = TRUE"
    ADOprimaryrs.Open TempStr, db, adOpenKeyset, adLockOptimistic, adCmdText
    Set grdDataGrid1.DataSource = ADOprimaryrs
    grdDataGrid1.ZOrder 0
    lblTop.Caption = grdDataGrid1.Caption
    Me.Caption = lblTop.Caption
    cmdCreate.Caption = "Create GL Entry"
Case 2
    TempStr = "SELECT [AR SALE Document #],[AR SALE Recur Type],[AR SALE Next Recur],[AR SALE Ext Document #],[AR SALE Document Type],[AR SALE Customer ID],[AR SALE Total] FROM [AR Sales] WHERE [AR SALE Recur Type]<> 'Never' "
    ADOprimaryrs.Open TempStr, db, adOpenKeyset, adLockOptimistic, adCmdText
    Set grdDataGrid2.DataSource = ADOprimaryrs
    grdDataGrid2.ZOrder 0
    lblTop.Caption = grdDataGrid2.Caption
    Me.Caption = lblTop.Caption
    cmdDate(0).Enabled = True
    cmdDate(1).Enabled = True
    cmdCreate.Caption = "Create Invoice"
    lblLabels(3).Enabled = True
    lblLabels(4).Enabled = True
    txtfields(0).Enabled = True
    txtfields(1).Enabled = True
    Label1.Enabled = True
Case 3
    TempStr = "SELECT [AP PO Document No],[AP PO Recurring YN],[AP PO Ext Document No],[AP PO Document Type],[AP PO Date],[AP PO Vendor ID],[AP PO Total Amount] FROM [AP Purchase] WHERE [AP PO Recurring YN] = TRUE"
    ADOprimaryrs.Open TempStr, db, adOpenKeyset, adLockOptimistic, adCmdText
    Set grdDataGrid3.DataSource = ADOprimaryrs
    grdDataGrid3.ZOrder 0
    lblTop.Caption = grdDataGrid3.Caption
    Me.Caption = lblTop.Caption
    cmdCreate.Caption = "Create Receiving"
Case 4
    TempStr = "SELECT [AP PAY ID],[AP PAY Recurring YN],[AP PAY Check No],[AP PAY Type],[AP PAY Transaction Date],[AP PAY Vendor No],[AP PAY Amount] FROM [AP Payment Header] WHERE [AP PAY Recurring YN] = TRUE"
    ADOprimaryrs.Open TempStr, db, adOpenKeyset, adLockOptimistic, adCmdText
    Set grdDataGrid4.DataSource = ADOprimaryrs
    grdDataGrid4.ZOrder 0
    lblTop.Caption = grdDataGrid4.Caption
    Me.Caption = lblTop.Caption
    cmdCreate.Caption = "Create Payment"
End Select
If ADOprimaryrs.RecordCount > 0 Then
    Me.Show
    GetTextColor Me
Else
    MsgBox "There is no " & lblTop.Caption & " transaction to publish"
    Unload Me
End If
End Sub

Private Sub cbPurchase_Click(Index As Integer)
If cbPurchase(15).Text <> "" Then
    grdDataGrid2.Columns(0).Text = cbPurchase(15).Text
    Select Case cbPurchase(15).Text
    Case "Never"
    Case "Monthly"
        grdDataGrid2.Columns(1) = DateAdd("m", 1, FormatDate(Now))
    Case "Quarterly"
        grdDataGrid2.Columns(1) = DateAdd("q", 1, FormatDate(Now))
    Case "Annually"
        grdDataGrid2.Columns(1) = DateAdd("yyyy", 1, FormatDate(Now))
    End Select
    ADOprimaryrs.Update
End If
cbPurchase(15).Visible = False
End Sub

Private Sub cbPurchase_LostFocus(Index As Integer)
   CheckCombo cbPurchase(Index), "[RECURR TYPE]", "[RECUR_TYPE]", db, True
   cbPurchase(15).Visible = False
End Sub

Private Sub cmdCreate_Click()
If ADOprimaryrs.RecordCount = 0 Then Exit Sub
Dim Success%
'db.BeginTrans

    Select Case lblTop.Caption
    Case "Recurring General Ledger"
        Set grdDataGrid1.DataSource = Nothing
            ADOprimaryrs.MoveFirst
            Do While Not ADOprimaryrs.EOF
              Success% = CloneGLEntry(ADOprimaryrs("GL TRANS Number"), db)
              ADOprimaryrs.MoveNext
            Loop
            ADOprimaryrs.MoveFirst
        Set grdDataGrid1.DataSource = ADOprimaryrs
    Case "Recurring Sales"
        Set grdDataGrid2.DataSource = Nothing
            ADOprimaryrs.MoveFirst
            Do While Not ADOprimaryrs.EOF
             Success% = CloneSales(ADOprimaryrs![AR SALE Document #], db)
             Select Case ADOprimaryrs("AR SALE Recur Type")
              Case "Monthly"
                ADOprimaryrs("AR SALE Next Recur") = DateAdd("m", 1, ADOprimaryrs("AR SALE Next Recur"))
              Case "Quarterly"
                ADOprimaryrs("AR SALE Next Recur") = DateAdd("q", 1, ADOprimaryrs("AR SALE Next Recur"))
              Case "Annually"
                ADOprimaryrs("AR SALE Next Recur") = DateAdd("yyyy", 1, ADOprimaryrs("AR SALE Next Recur"))
              End Select
              ADOprimaryrs.Update
              ADOprimaryrs.MoveNext
            Loop
            ADOprimaryrs.MoveFirst
        Set grdDataGrid2.DataSource = ADOprimaryrs
    Case "Recurring Purchases"
        Set grdDataGrid3.DataSource = Nothing
            ADOprimaryrs.MoveFirst
            Do While Not ADOprimaryrs.EOF
              Success% = ClonePurchase(ADOprimaryrs("AP PO Document No"), db)
              ADOprimaryrs.MoveNext
            Loop
            ADOprimaryrs.MoveFirst
        Set grdDataGrid3.DataSource = ADOprimaryrs
    Case "Recurring Payments"
        Set grdDataGrid4.DataSource = Nothing
            ADOprimaryrs.MoveFirst
            Do While Not ADOprimaryrs.EOF
              Success% = ClonePayment(ADOprimaryrs("AP PAY ID"), db)
              ADOprimaryrs.MoveNext
            Loop
            ADOprimaryrs.MoveFirst
        Set grdDataGrid4.DataSource = ADOprimaryrs
    End Select
    
    ProcessDoneMusic "Done"
'db.CommitTrans
End Sub

Private Sub cmdDate_Click(Index As Integer)
Select Case Index
Case 0
    Menu_Calendar.WhoCallMe True, 1302
Case 1
    Menu_Calendar.WhoCallMe True, 1640
End Select
End Sub

Private Sub cmdSearch_Click()
If ADOprimaryrs Is Nothing Then
Else
    If ADOprimaryrs.RecordCount = 0 Then Exit Sub
        Select Case lblTop.Caption
        Case "Recurring General Ledger"
            SearchRECORD ADOprimaryrs, grdDataGrid1, txtfields(2).Text, lblfields(0).Caption, WhichField, "GL TRANS Document #"
        Case "Recurring Sales"
            SearchRECORD ADOprimaryrs, grdDataGrid2, txtfields(2).Text, lblfields(0).Caption, WhichField, "AR SALE Ext Document #"
        Case "Recurring Purchases"
            SearchRECORD ADOprimaryrs, grdDataGrid3, txtfields(2).Text, lblfields(0).Caption, WhichField, "AP PO Ext Document No"
        Case "Recurring Payments"
            SearchRECORD ADOprimaryrs, grdDataGrid4, txtfields(2).Text, lblfields(0).Caption, WhichField, "AP PAY Check No"
        End Select
    ProcessDoneMusic "Done"
End If
End Sub

Private Sub Form_Load()
  Me.Width = 13095
  Me.Height = 5745
  loadCombo
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  If fMainForm.WindowState = 1 Then Exit Sub
  If Me.WindowState = 0 Then
  ElseIf Me.WindowState = 2 Then
    GoTo SkipResize
  Else
    Exit Sub
  End If
  Me.Width = 13095
  Me.Height = 5745

SkipResize:
  frPrimary.Left = (Me.ScaleWidth - frPrimary.Width) / 2
  lblTop.Left = frPrimary.Left
  lblTop.Width = frPrimary.Width
  frPrimary.Top = (Me.ScaleHeight - frPrimary.Height) / 2 + 230
  grdDataGrid3.Left = grdDataGrid1.Left
  grdDataGrid3.Top = grdDataGrid1.Top
  grdDataGrid3.Height = grdDataGrid1.Height
  grdDataGrid4.Left = grdDataGrid1.Left
  grdDataGrid4.Top = grdDataGrid1.Top
  grdDataGrid4.Height = grdDataGrid1.Height
  grdDataGrid2.Left = grdDataGrid1.Left
  grdDataGrid2.Top = grdDataGrid1.Top
  grdDataGrid2.Height = grdDataGrid1.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If ADOprimaryrs.RecordCount > 0 Then ADOprimaryrs.Update
    ADOprimaryrs.Close
    Set ADOprimaryrs = Nothing
    db.Close
    Set db = Nothing
    Set frm_Recurring = Nothing
End Sub


Private Sub grdDataGrid1_ButtonClick(ByVal ColIndex As Integer)
On Error GoTo Error_ButtClick
If ADOprimaryrs.RecordCount = 0 Then Exit Sub
   Select Case ColIndex
   Case 0
      If grdDataGrid1.Columns(0).Text = "No" Then
         grdDataGrid1.Columns(0).Text = "Yes"
      Else
         grdDataGrid1.Columns(0).Text = "No"
      End If
         SendKeys ("{ENTER}")
         SendKeys ("{down}")
         SendKeys ("{up}")
End Select
Exit Sub
Error_ButtClick:
    MsgBox "Please click the Table box before clicking the button"
End Sub

Private Sub grdDatagrid1_HeadClick(ByVal ColIndex As Integer)
If ADOprimaryrs.RecordCount = 0 Then Exit Sub
    lblfields(0) = grdDataGrid1.Columns(ColIndex).Caption
    WhichField = grdDataGrid1.Columns(ColIndex).DataField
    ADOprimaryrs.Close
    Set ADOprimaryrs = Nothing
    Set ADOprimaryrs = New ADODB.Recordset
    ADOprimaryrs.Open TempStr & " ORDER BY [" & WhichField & "]", db, adOpenKeyset, adLockOptimistic, adCmdText
    Set grdDataGrid1.DataSource = ADOprimaryrs

End Sub

Private Sub grdDataGrid2_ButtonClick(ByVal ColIndex As Integer)
On Error GoTo Error_ButtClick
If ADOprimaryrs.RecordCount = 0 Then Exit Sub
Select Case ColIndex
Case 0
    cbPurchase(15).Width = grdDataGrid2.Columns(ColIndex).Width
    cbPurchase(15).Top = (grdDataGrid2.RowHeight * (grdDataGrid2.Row + 2)) + grdDataGrid2.Top - 30
    cbPurchase(15).Left = grdDataGrid2.Columns(ColIndex).Left + grdDataGrid2.Left
    cbPurchase(15).Visible = True
    cbPurchase(15).ZOrder 0
    cbPurchase(15).SetFocus
Case 1
    Menu_Calendar.WhoCallMe True, 1660
    SendKeys ("{ENTER}")
    SendKeys ("{down}")
    SendKeys ("{up}")
End Select
Exit Sub
Error_ButtClick:
    MsgBox "Please click the Table box before clicking the button"
End Sub

Private Sub grdDatagrid2_HeadClick(ByVal ColIndex As Integer)
If ADOprimaryrs.RecordCount = 0 Then Exit Sub
    lblfields(0) = grdDataGrid2.Columns(ColIndex).Caption
    WhichField = grdDataGrid2.Columns(ColIndex).DataField
    ADOprimaryrs.Close
    Set ADOprimaryrs = Nothing
    Set ADOprimaryrs = New ADODB.Recordset
    ADOprimaryrs.Open TempStr & " ORDER BY [" & WhichField & "]", db, adOpenKeyset, adLockOptimistic, adCmdText
    Set grdDataGrid2.DataSource = ADOprimaryrs

End Sub

Private Sub grdDataGrid3_ButtonClick(ByVal ColIndex As Integer)
On Error GoTo Error_ButtClick
If ADOprimaryrs.RecordCount = 0 Then Exit Sub
   
   Select Case ColIndex
   Case 0
      If grdDataGrid3.Columns(0).Text = "No" Then
         grdDataGrid3.Columns(0).Text = "Yes"
      Else
         grdDataGrid3.Columns(0).Text = "No"
      End If
         SendKeys ("{ENTER}")
         SendKeys ("{down}")
         SendKeys ("{up}")
   End Select
Exit Sub
Error_ButtClick:
    MsgBox "Please click the Table box before clicking the button"
End Sub

Private Sub grdDatagrid3_HeadClick(ByVal ColIndex As Integer)
If ADOprimaryrs.RecordCount = 0 Then Exit Sub
    lblfields(0) = grdDataGrid3.Columns(ColIndex).Caption
    WhichField = grdDataGrid3.Columns(ColIndex).DataField
    ADOprimaryrs.Close
    Set ADOprimaryrs = Nothing
    Set ADOprimaryrs = New ADODB.Recordset
    ADOprimaryrs.Open TempStr & " ORDER BY [" & WhichField & "]", db, adOpenKeyset, adLockOptimistic, adCmdText
    Set grdDataGrid3.DataSource = ADOprimaryrs

End Sub

Private Sub grdDataGrid4_ButtonClick(ByVal ColIndex As Integer)
On Error GoTo Error_ButtClick
If ADOprimaryrs.RecordCount = 0 Then Exit Sub
   Select Case ColIndex
   Case 0
      If grdDataGrid4.Columns(0).Text = "No" Then
         grdDataGrid4.Columns(0).Text = "Yes"
      Else
         grdDataGrid41.Columns(0).Text = "No"
      End If
         SendKeys ("{ENTER}")
         SendKeys ("{down}")
         SendKeys ("{up}")
   End Select
Exit Sub
Error_ButtClick:
    MsgBox "Please click the Table box before clicking the button"
End Sub

Private Sub grdDataGrid4_HeadClick(ByVal ColIndex As Integer)
If ADOprimaryrs.RecordCount = 0 Then Exit Sub
    lblfields(0) = grdDataGrid4.Columns(ColIndex).Caption
    WhichField = grdDataGrid4.Columns(ColIndex).DataField
    ADOprimaryrs.Close
    Set ADOprimaryrs = Nothing
    Set ADOprimaryrs = New ADODB.Recordset
    ADOprimaryrs.Open TempStr & " ORDER BY [" & WhichField & "]", db, adOpenKeyset, adLockOptimistic, adCmdText
    Set grdDataGrid4.DataSource = ADOprimaryrs

End Sub
Private Sub loadCombo()
    ComboInit cbPurchase(15), lblfields(15), "SELECT [RECURR TYPE] FROM [RECUR_TYPE]", db
End Sub
