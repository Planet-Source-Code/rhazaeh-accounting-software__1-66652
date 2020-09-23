VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_AP_Vendor 
   Caption         =   "Vendor Data"
   ClientHeight    =   8370
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   10635
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8370
   ScaleWidth      =   10635
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   10635
      TabIndex        =   46
      Top             =   7770
      Width           =   10635
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   300
         Left            =   4320
         TabIndex        =   40
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   300
         Left            =   3240
         TabIndex        =   39
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   300
         Left            =   2160
         TabIndex        =   38
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   300
         Left            =   1080
         TabIndex        =   37
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   300
         Left            =   0
         TabIndex        =   36
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
      ScaleWidth      =   10635
      TabIndex        =   0
      Top             =   8070
      Width           =   10635
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Picture         =   "frm_AP_Vendor.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Picture         =   "frm_AP_Vendor.frx":0342
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Picture         =   "frm_AP_Vendor.frx":0684
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Picture         =   "frm_AP_Vendor.frx":09C6
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   45
         Top             =   0
         Width           =   3480
      End
   End
   Begin VB.Frame frShowAll 
      Height          =   7215
      Left            =   0
      TabIndex        =   112
      Top             =   480
      Visible         =   0   'False
      Width           =   10575
      Begin VB.CommandButton cmdBack 
         Caption         =   "Back"
         Height          =   795
         Left            =   8400
         Picture         =   "frm_AP_Vendor.frx":0D08
         Style           =   1  'Graphical
         TabIndex        =   118
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         Height          =   795
         Left            =   9480
         Picture         =   "frm_AP_Vendor.frx":1012
         Style           =   1  'Graphical
         TabIndex        =   117
         Top             =   240
         Width           =   975
      End
      Begin VB.Frame Frame5 
         Height          =   975
         Left            =   5760
         TabIndex        =   113
         Top             =   120
         Width           =   2535
         Begin VB.CommandButton Command4 
            Height          =   540
            Left            =   1680
            Picture         =   "frm_AP_Vendor.frx":131C
            Style           =   1  'Graphical
            TabIndex        =   115
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
            Index           =   45
            Left            =   240
            TabIndex        =   114
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "ID"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   34
            Left            =   240
            TabIndex        =   116
            Top             =   240
            Width           =   1335
         End
      End
      Begin MSDataGridLib.DataGrid grdDataGrid 
         Bindings        =   "frm_AP_Vendor.frx":1626
         Height          =   5895
         Left            =   120
         TabIndex        =   119
         Top             =   1200
         Width           =   10305
         _ExtentX        =   18177
         _ExtentY        =   10398
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
         Caption         =   "Vendor Data"
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "AP VEN ID"
            Caption         =   "Vendor ID"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "h:mm:ss AMPM"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "AP VEN Name"
            Caption         =   "Name"
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
            DataField       =   "AP VEN Contact"
            Caption         =   "Contact"
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
            DataField       =   "AP VEN Phone"
            Caption         =   "Phone"
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
            DataField       =   "AP VEN Phone Ext"
            Caption         =   "Ext"
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
            DataField       =   "AP VEN Payments YTD"
            Caption         =   "Payment YTD"
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
            DataField       =   "AP VEN Notes"
            Caption         =   "Notes"
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
            MarqueeStyle    =   4
            BeginProperty Column00 
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1904.882
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1395.213
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   599.811
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               ColumnWidth     =   1110.047
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   2009.764
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame frPrimary 
      Height          =   7215
      Left            =   0
      TabIndex        =   47
      Top             =   480
      Width           =   10575
      Begin VB.TextBox txtfields 
         DataField       =   "AP VEN Credit Limit"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00;(""$""#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   16
         Left            =   4800
         TabIndex        =   94
         Top             =   6000
         Width           =   1455
      End
      Begin VB.CheckBox chkFields 
         Alignment       =   1  'Right Justify
         Caption         =   "1099 Vendor"
         DataField       =   "AP VEN 1099"
         Height          =   255
         Index           =   1
         Left            =   4800
         TabIndex        =   93
         Top             =   6360
         Width           =   1455
      End
      Begin VB.CommandButton Tranbut 
         Caption         =   "Purchase"
         Height          =   795
         Left            =   7680
         Picture         =   "frm_AP_Vendor.frx":1641
         Style           =   1  'Graphical
         TabIndex        =   92
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton FinBut 
         Caption         =   "Financials"
         Height          =   795
         Left            =   8640
         Picture         =   "frm_AP_Vendor.frx":194B
         Style           =   1  'Graphical
         TabIndex        =   91
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cdmShowAll 
         Caption         =   "List All"
         Height          =   795
         Left            =   9600
         Picture         =   "frm_AP_Vendor.frx":1C55
         Style           =   1  'Graphical
         TabIndex        =   90
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtfields 
         DataField       =   "AP VEN Tax ID"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   1
         Left            =   4800
         TabIndex        =   88
         Top             =   6720
         Width           =   1455
      End
      Begin VB.CommandButton cmdGL 
         Height          =   285
         Left            =   3240
         Picture         =   "frm_AP_Vendor.frx":1F5F
         Style           =   1  'Graphical
         TabIndex        =   87
         Top             =   6000
         Width           =   375
      End
      Begin VB.CheckBox chkFields 
         Alignment       =   1  'Right Justify
         Caption         =   "Inactive Vendor"
         DataField       =   "AP VEN Inactive YN"
         DataSource      =   "adoPrimaryRS"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   5040
         TabIndex        =   86
         Top             =   720
         Width           =   1455
      End
      Begin VB.CheckBox chkFields 
         Alignment       =   1  'Right Justify
         Caption         =   "Payment Hold"
         DataField       =   "AP VEN Payment Hold"
         DataSource      =   "adoPrimaryRS"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   16
         Left            =   5040
         TabIndex        =   85
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtfields 
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   34
         Left            =   1680
         TabIndex        =   83
         Top             =   480
         Width           =   1575
      End
      Begin VB.Frame Frame2 
         Height          =   4815
         Left            =   5400
         TabIndex        =   68
         Top             =   1080
         Width           =   5055
         Begin VB.CommandButton Command1 
            Caption         =   "Default >>"
            Height          =   375
            Left            =   120
            TabIndex        =   84
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txtfields 
            DataField       =   "AP VEN Custom 3"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   18
            Left            =   1440
            TabIndex        =   18
            Top             =   1080
            Width           =   3375
         End
         Begin VB.TextBox txtfields 
            DataField       =   "AP VEN Custom 4"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   26
            Left            =   1440
            TabIndex        =   26
            Top             =   3240
            Width           =   3375
         End
         Begin VB.TextBox txtfields 
            DataField       =   "AP VEN Remit Phone Ext"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   28
            Left            =   3720
            TabIndex        =   28
            Top             =   3600
            Width           =   1095
         End
         Begin VB.TextBox txtfields 
            DataField       =   "AP VEN Remit Phone"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   27
            Left            =   1440
            TabIndex        =   27
            Top             =   3600
            Width           =   1695
         End
         Begin VB.TextBox txtfields 
            DataField       =   "AP VEN Remit Name"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   17
            Left            =   1440
            TabIndex        =   17
            Top             =   720
            Width           =   3375
         End
         Begin VB.TextBox txtfields 
            DataField       =   "AP VEN Remit Fax"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   29
            Left            =   1440
            TabIndex        =   29
            Top             =   3960
            Width           =   1695
         End
         Begin VB.TextBox txtfields 
            DataField       =   "AP VEN Remit Country"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   24
            Left            =   1440
            TabIndex        =   24
            Top             =   2520
            Width           =   1575
         End
         Begin VB.TextBox txtfields 
            DataField       =   "AP VEN Remit Contact"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   25
            Left            =   1440
            TabIndex        =   25
            Top             =   2880
            Width           =   3375
         End
         Begin VB.TextBox txtfields 
            DataField       =   "AP VEN Remit City"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   21
            Left            =   1440
            TabIndex        =   21
            Top             =   2160
            Width           =   975
         End
         Begin VB.TextBox txtfields 
            DataField       =   "AP VEN Remit Address 2"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   20
            Left            =   1440
            TabIndex        =   20
            Top             =   1800
            Width           =   3375
         End
         Begin VB.TextBox txtfields 
            DataField       =   "AP VEN Remit Address 1"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   19
            Left            =   1440
            TabIndex        =   19
            Top             =   1440
            Width           =   3375
         End
         Begin VB.TextBox txtfields 
            DataField       =   "AP VEN Remit State"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   22
            Left            =   3000
            TabIndex        =   22
            Top             =   2160
            Width           =   615
         End
         Begin VB.TextBox txtfields 
            DataField       =   "AP VEN Remit Postal"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   23
            Left            =   4080
            TabIndex        =   23
            Top             =   2160
            Width           =   735
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Web Page Add:  "
            DataSource      =   "adoPrimaryRS"
            Height          =   255
            Index           =   24
            Left            =   120
            TabIndex        =   82
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "E-mail Address:  "
            DataSource      =   "adoPrimaryRS"
            Height          =   255
            Index           =   19
            Left            =   120
            TabIndex        =   81
            Top             =   3240
            Width           =   1335
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Ext:  "
            DataSource      =   "adoPrimaryRS"
            Height          =   255
            Index           =   31
            Left            =   3120
            TabIndex        =   78
            Top             =   3600
            Width           =   615
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Telephone No:  "
            DataSource      =   "adoPrimaryRS"
            Height          =   255
            Index           =   30
            Left            =   120
            TabIndex        =   77
            Top             =   3600
            Width           =   1335
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Remit Name:  "
            DataSource      =   "adoPrimaryRS"
            Height          =   255
            Index           =   29
            Left            =   120
            TabIndex        =   76
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Fax Number:  "
            DataSource      =   "adoPrimaryRS"
            Height          =   255
            Index           =   28
            Left            =   120
            TabIndex        =   75
            Top             =   3960
            Width           =   1335
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Country:  "
            DataSource      =   "adoPrimaryRS"
            Height          =   255
            Index           =   27
            Left            =   120
            TabIndex        =   74
            Top             =   2520
            Width           =   1335
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Contact Person:  "
            DataSource      =   "adoPrimaryRS"
            Height          =   255
            Index           =   26
            Left            =   120
            TabIndex        =   73
            Top             =   2880
            Width           =   1335
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "City:  "
            DataSource      =   "adoPrimaryRS"
            Height          =   255
            Index           =   25
            Left            =   120
            TabIndex        =   72
            Top             =   2160
            Width           =   1335
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Address:  "
            DataSource      =   "adoPrimaryRS"
            Height          =   255
            Index           =   23
            Left            =   120
            TabIndex        =   71
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "State:  "
            DataSource      =   "adoPrimaryRS"
            Height          =   255
            Index           =   1
            Left            =   2400
            TabIndex        =   70
            Top             =   2160
            Width           =   615
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Zip:  "
            DataSource      =   "adoPrimaryRS"
            Height          =   255
            Index           =   0
            Left            =   3600
            TabIndex        =   69
            Top             =   2160
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         Height          =   4815
         Left            =   120
         TabIndex        =   55
         Top             =   1080
         Width           =   5175
         Begin VB.TextBox txtfields 
            DataField       =   "AP VEN Custom 2"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   11
            Left            =   1560
            TabIndex        =   11
            Top             =   3240
            Width           =   3375
         End
         Begin VB.TextBox txtfields 
            DataField       =   "AP VEN Custom 1"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   3
            Left            =   1560
            TabIndex        =   3
            Top             =   1080
            Width           =   3375
         End
         Begin VB.TextBox txtfields 
            DataField       =   "AP VEN Postal"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   8
            Left            =   4200
            TabIndex        =   8
            Top             =   2160
            Width           =   735
         End
         Begin VB.TextBox txtfields 
            DataField       =   "AP VEN Phone Ext"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   13
            Left            =   3720
            TabIndex        =   13
            Top             =   3600
            Width           =   1215
         End
         Begin VB.TextBox txtfields 
            DataField       =   "AP VEN Phone"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   12
            Left            =   1560
            TabIndex        =   12
            Top             =   3600
            Width           =   1575
         End
         Begin VB.TextBox txtfields 
            DataField       =   "AP VEN Name"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   2
            Left            =   1560
            TabIndex        =   2
            Top             =   720
            Width           =   3375
         End
         Begin VB.TextBox txtfields 
            DataField       =   "AP VEN Fax"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   14
            Left            =   1560
            TabIndex        =   14
            Top             =   3960
            Width           =   1575
         End
         Begin VB.TextBox txtfields 
            DataField       =   "AP VEN Department"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   15
            Left            =   1560
            TabIndex        =   15
            Top             =   4320
            Width           =   975
         End
         Begin VB.TextBox txtfields 
            DataField       =   "AP VEN Default Department"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   121
            Left            =   3960
            TabIndex        =   16
            Top             =   4320
            Width           =   975
         End
         Begin VB.TextBox txtfields 
            DataField       =   "AP VEN Country"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   9
            Left            =   1560
            TabIndex        =   9
            Top             =   2520
            Width           =   1575
         End
         Begin VB.TextBox txtfields 
            DataField       =   "AP VEN Contact"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   10
            Left            =   1560
            TabIndex        =   10
            Top             =   2880
            Width           =   3375
         End
         Begin VB.TextBox txtfields 
            DataField       =   "AP VEN City"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   6
            Left            =   1560
            TabIndex        =   6
            Top             =   2160
            Width           =   975
         End
         Begin VB.TextBox txtfields 
            DataField       =   "AP VEN Address 2"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   5
            Left            =   1560
            TabIndex        =   5
            Top             =   1800
            Width           =   3375
         End
         Begin VB.TextBox txtfields 
            DataField       =   "AP VEN Address 1"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   4
            Left            =   1560
            TabIndex        =   4
            Top             =   1440
            Width           =   3375
         End
         Begin VB.TextBox txtfields 
            DataField       =   "AP VEN State"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   7
            Left            =   3120
            TabIndex        =   7
            Top             =   2160
            Width           =   615
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "E-mail Address:  "
            DataSource      =   "adoPrimaryRS"
            Height          =   255
            Index           =   16
            Left            =   120
            TabIndex        =   80
            Top             =   3240
            Width           =   1455
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Web Page Add:  "
            DataSource      =   "adoPrimaryRS"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   79
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Zip:  "
            DataSource      =   "adoPrimaryRS"
            Height          =   255
            Index           =   22
            Left            =   3720
            TabIndex        =   67
            Top             =   2160
            Width           =   495
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Ext:  "
            DataSource      =   "adoPrimaryRS"
            Height          =   255
            Index           =   21
            Left            =   3120
            TabIndex        =   66
            Top             =   3600
            Width           =   615
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Telephone No:  "
            DataSource      =   "adoPrimaryRS"
            Height          =   255
            Index           =   20
            Left            =   120
            TabIndex        =   65
            Top             =   3600
            Width           =   1455
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Vendor Name:  "
            DataSource      =   "adoPrimaryRS"
            Height          =   255
            Index           =   17
            Left            =   120
            TabIndex        =   64
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Fax Number:  "
            DataSource      =   "adoPrimaryRS"
            Height          =   255
            Index           =   14
            Left            =   120
            TabIndex        =   63
            Top             =   3960
            Width           =   1455
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Department:  "
            DataSource      =   "adoPrimaryRS"
            Height          =   255
            Index           =   13
            Left            =   120
            TabIndex        =   62
            Top             =   4320
            Width           =   1455
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Default Dept:  "
            DataSource      =   "adoPrimaryRS"
            Height          =   255
            Index           =   10
            Left            =   2520
            TabIndex        =   61
            Top             =   4320
            Width           =   1455
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Country:  "
            DataSource      =   "adoPrimaryRS"
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   60
            Top             =   2520
            Width           =   1455
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Contact Person:  "
            DataSource      =   "adoPrimaryRS"
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   59
            Top             =   2880
            Width           =   1455
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "City:  "
            DataSource      =   "adoPrimaryRS"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   58
            Top             =   2160
            Width           =   1455
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Address:  "
            DataSource      =   "adoPrimaryRS"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   57
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "State:  "
            DataSource      =   "adoPrimaryRS"
            Height          =   255
            Index           =   2
            Left            =   2520
            TabIndex        =   56
            Top             =   2160
            Width           =   615
         End
      End
      Begin VB.TextBox txtfields 
         DataField       =   "AP VEN Notes"
         DataSource      =   "adoPrimaryRS"
         Height          =   1125
         Index           =   31
         Left            =   6960
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   35
         Top             =   6000
         Width           =   3495
      End
      Begin VB.TextBox txtfields 
         DataField       =   "AP VEN Default GL"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   30
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   6000
         Width           =   1575
      End
      Begin VB.TextBox txtfields 
         DataField       =   "AP VEN ID"
         DataSource      =   "adoPrimaryRS"
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   1
         Top             =   480
         Width           =   1575
      End
      Begin VB.CommandButton btVenID 
         Height          =   285
         Left            =   3240
         Picture         =   "frm_AP_Vendor.frx":2269
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   480
         Width           =   375
      End
      Begin VB.ComboBox cbfields 
         DataField       =   "AP VEN Type"
         Height          =   315
         Index           =   1
         ItemData        =   "frm_AP_Vendor.frx":2573
         Left            =   1680
         List            =   "frm_AP_Vendor.frx":2575
         TabIndex        =   31
         Top             =   6360
         Width           =   1575
      End
      Begin VB.ComboBox cbfields 
         DataField       =   "AP VEN Default Terms"
         Height          =   315
         Index           =   2
         Left            =   1680
         TabIndex        =   33
         Top             =   6720
         Width           =   1575
      End
      Begin VB.CommandButton btcbRefresh 
         Height          =   300
         Index           =   1
         Left            =   3240
         Picture         =   "frm_AP_Vendor.frx":2577
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Update the Ship Via"
         Top             =   6360
         Width           =   375
      End
      Begin VB.CommandButton btcbRefresh 
         Height          =   300
         Index           =   2
         Left            =   3240
         Picture         =   "frm_AP_Vendor.frx":2881
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Update the Ship Via"
         Top             =   6720
         Width           =   375
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Credit Limit:  "
         DataSource      =   "adoPrimaryRS"
         Height          =   255
         Index           =   4
         Left            =   3840
         TabIndex        =   95
         Top             =   6000
         Width           =   975
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Tax ID:  "
         DataSource      =   "adoPrimaryRS"
         Height          =   255
         Index           =   3
         Left            =   3840
         TabIndex        =   89
         Top             =   6720
         Width           =   975
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Notes:  "
         DataSource      =   "adoPrimaryRS"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   18
         Left            =   6360
         TabIndex        =   53
         Top             =   6000
         Width           =   615
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Vendor ID:  "
         DataSource      =   "adoPrimaryRS"
         Height          =   255
         Index           =   15
         Left            =   240
         TabIndex        =   52
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label lblFields 
         Alignment       =   1  'Right Justify
         Caption         =   "Payment Terms   "
         DataSource      =   "adoPrimaryRS"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   51
         Top             =   6720
         Width           =   1455
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Default GL   "
         DataSource      =   "adoPrimaryRS"
         Height          =   255
         Index           =   11
         Left            =   240
         TabIndex        =   50
         Top             =   6000
         Width           =   1455
      End
      Begin VB.Label lblFields 
         Alignment       =   1  'Right Justify
         Caption         =   "Vendor Type   "
         DataSource      =   "adoPrimaryRS"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   49
         Top             =   6360
         Width           =   1455
      End
   End
   Begin VB.Frame AP_VendorTrans 
      Height          =   7215
      Left            =   0
      TabIndex        =   96
      Top             =   480
      Visible         =   0   'False
      Width           =   10575
      Begin VB.Frame Frame4 
         Height          =   1335
         Left            =   5280
         TabIndex        =   107
         Top             =   240
         Width           =   2535
         Begin VB.CommandButton cmdSearch 
            Height          =   540
            Left            =   1680
            Picture         =   "frm_AP_Vendor.frx":2B8B
            Style           =   1  'Graphical
            TabIndex        =   109
            Top             =   480
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
            Index           =   35
            Left            =   240
            TabIndex        =   108
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Document No."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   33
            Left            =   240
            TabIndex        =   110
            Top             =   480
            Width           =   1335
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1335
         Left            =   120
         TabIndex        =   98
         Top             =   240
         Width           =   5055
         Begin VB.CommandButton cmdShow 
            Caption         =   "&Execute"
            Height          =   735
            Left            =   3000
            Picture         =   "frm_AP_Vendor.frx":2E95
            Style           =   1  'Graphical
            TabIndex        =   104
            Top             =   360
            Width           =   855
         End
         Begin VB.CommandButton cmdDate 
            Height          =   285
            Index           =   1
            Left            =   2400
            Picture         =   "frm_AP_Vendor.frx":319F
            Style           =   1  'Graphical
            TabIndex        =   103
            Top             =   840
            Width           =   375
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
            Index           =   33
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   102
            Top             =   840
            Width           =   1455
         End
         Begin VB.CommandButton cmdDate 
            Height          =   285
            Index           =   0
            Left            =   2400
            Picture         =   "frm_AP_Vendor.frx":34A9
            Style           =   1  'Graphical
            TabIndex        =   101
            Top             =   360
            Width           =   375
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
            Index           =   32
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   100
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton Command2 
            Caption         =   "A&ll"
            Height          =   735
            Left            =   4080
            Picture         =   "frm_AP_Vendor.frx":37B3
            Style           =   1  'Graphical
            TabIndex        =   99
            Top             =   360
            Width           =   855
         End
         Begin VB.Label lblLabels 
            Caption         =   "End Date:"
            Height          =   255
            Index           =   32
            Left            =   120
            TabIndex        =   106
            Top             =   840
            Width           =   975
         End
         Begin VB.Label lblLabels 
            Caption         =   "Start Date:"
            Height          =   255
            Index           =   12
            Left            =   120
            TabIndex        =   105
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.CommandButton cmdBack1 
         Caption         =   "&Back"
         Height          =   855
         Left            =   9120
         Picture         =   "frm_AP_Vendor.frx":3ABD
         Style           =   1  'Graphical
         TabIndex        =   97
         Top             =   600
         Width           =   975
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Height          =   5415
         Left            =   120
         TabIndex        =   111
         Top             =   1680
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   9551
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
         Caption         =   "Vendor Transaction"
         ColumnCount     =   8
         BeginProperty Column00 
            DataField       =   "AP PO Posted YN"
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
         BeginProperty Column01 
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
         BeginProperty Column02 
            DataField       =   "AP PO Document Type"
            Caption         =   "Doc. Type"
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
            DataField       =   "AP PO Vendor Invoice No"
            Caption         =   "Vendor Invoice"
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
            DataField       =   "AP PO Ext Document No"
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
         BeginProperty Column05 
            DataField       =   "AP PO Payment Terms"
            Caption         =   "Payment Terms"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   """$""#,##0.00;(""$""#,##0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
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
         BeginProperty Column07 
            DataField       =   "AP PO Total Amount"
            Caption         =   "Purchase Total"
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
               ColumnWidth     =   764.787
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1244.976
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1230.236
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1349.858
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1244.976
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1395.213
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1244.976
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1244.976
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame AP_VENDORfinancials 
      Height          =   7215
      Left            =   0
      TabIndex        =   120
      Top             =   480
      Visible         =   0   'False
      Width           =   10575
      Begin VB.CommandButton cmdAgingDetails 
         Caption         =   "Aging Detail"
         Height          =   375
         Left            =   120
         TabIndex        =   121
         Top             =   4200
         Width           =   1335
      End
      Begin VB.CommandButton CmdPrtState 
         Caption         =   "Print Statement"
         Height          =   375
         Left            =   120
         TabIndex        =   122
         Top             =   6360
         Width           =   1335
      End
      Begin VB.CommandButton cdmAgeReceivable 
         Caption         =   "Age Receivable"
         Height          =   375
         Left            =   120
         TabIndex        =   123
         Top             =   6000
         Width           =   1335
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Calcula&te"
         Height          =   855
         Left            =   9240
         Picture         =   "frm_AP_Vendor.frx":3DC7
         Style           =   1  'Graphical
         TabIndex        =   144
         Top             =   3240
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Back"
         Height          =   855
         Left            =   9240
         Picture         =   "frm_AP_Vendor.frx":40D1
         Style           =   1  'Graphical
         TabIndex        =   143
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox txtfields 
         DataField       =   "AP VEN ID"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   57
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   142
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtfields 
         DataField       =   "AP VEN Name"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   56
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   141
         Top             =   960
         Width           =   3375
      End
      Begin VB.TextBox txtfields 
         Alignment       =   1  'Right Justify
         DataField       =   "AP VEN Financial Period 1"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00;(""$""#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   55
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   140
         Top             =   1800
         Width           =   1815
      End
      Begin VB.TextBox txtfields 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF00&
         DataField       =   "AP VEN Financial Period 2"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00;(""$""#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   54
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   139
         Top             =   1800
         Width           =   1815
      End
      Begin VB.TextBox txtfields 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000FFFF&
         DataField       =   "AP VEN Financial Period 3"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00;(""$""#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   53
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   138
         Top             =   1800
         Width           =   1815
      End
      Begin VB.TextBox txtfields 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         DataField       =   "AP VEN Financial Period 4"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00;(""$""#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   52
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   137
         Top             =   1800
         Width           =   1815
      End
      Begin VB.TextBox txtfields 
         Alignment       =   1  'Right Justify
         DataField       =   "AP VEN Financial Total"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00;(""$""#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   51
         Left            =   8160
         Locked          =   -1  'True
         TabIndex        =   136
         Top             =   1800
         Width           =   1815
      End
      Begin VB.TextBox txtfields 
         Alignment       =   1  'Right Justify
         DataField       =   "AP VEN Purchase Last Year"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00;(""$""#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   49
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   135
         Top             =   2640
         Width           =   1815
      End
      Begin VB.TextBox txtfields 
         Alignment       =   1  'Right Justify
         DataField       =   "AP VEN Purchase Lifetime"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00;(""$""#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   48
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   134
         Top             =   2640
         Width           =   1815
      End
      Begin VB.TextBox txtfields 
         Alignment       =   1  'Right Justify
         DataField       =   "AP VEN Purchase YTD"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00;(""$""#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   47
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   133
         Top             =   2640
         Width           =   1815
      End
      Begin VB.TextBox txtfields 
         Alignment       =   2  'Center
         DataField       =   "AP VEN Purchase Number Last Year"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """$""#,##0.00;(""$""#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   46
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   132
         Top             =   3360
         Width           =   1815
      End
      Begin VB.TextBox txtfields 
         Alignment       =   2  'Center
         DataField       =   "AP VEN Purchase Number Lifetime"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """$""#,##0.00;(""$""#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   36
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   131
         Top             =   3360
         Width           =   1815
      End
      Begin VB.TextBox txtfields 
         Alignment       =   2  'Center
         DataField       =   "AP VEN Purchase Number YTD"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """$""#,##0.00;(""$""#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   37
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   130
         Top             =   3360
         Width           =   1815
      End
      Begin VB.TextBox txtfields 
         Alignment       =   1  'Right Justify
         DataField       =   "AP VEN Payments Last Year"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00;(""$""#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   38
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   129
         Top             =   3000
         Width           =   1815
      End
      Begin VB.TextBox txtfields 
         Alignment       =   1  'Right Justify
         DataField       =   "AP VEN Payments Lifetime"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00;(""$""#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   39
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   128
         Top             =   3000
         Width           =   1815
      End
      Begin VB.TextBox txtfields 
         Alignment       =   1  'Right Justify
         DataField       =   "AP VEN Payments YTD"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00;(""$""#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   40
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   127
         Top             =   3000
         Width           =   1815
      End
      Begin VB.TextBox txtfields 
         Alignment       =   2  'Center
         DataField       =   "AP VEN Payment Number Last Year"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """$""#,##0.00;(""$""#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   41
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   126
         Top             =   3720
         Width           =   1815
      End
      Begin VB.TextBox txtfields 
         Alignment       =   2  'Center
         DataField       =   "AP VEN Payment Number Lifetime"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """$""#,##0.00;(""$""#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   42
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   125
         Top             =   3720
         Width           =   1815
      End
      Begin VB.TextBox txtfields 
         Alignment       =   2  'Center
         DataField       =   "AP VEN Payment Number YTD"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """$""#,##0.00;(""$""#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   43
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   124
         Top             =   3720
         Width           =   1815
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2895
         Left            =   2400
         TabIndex        =   145
         Top             =   4200
         Visible         =   0   'False
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   5106
         _Version        =   393216
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
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "AGE PO Doc Ext No"
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
         BeginProperty Column01 
            DataField       =   "AGE Start Date"
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
            DataField       =   "AGE Orig Amount"
            Caption         =   "Orig. Amount"
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
            DataField       =   "AGE PEriod 1"
            Caption         =   "Period 1"
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
            DataField       =   "AGE Period 2"
            Caption         =   "Period 2"
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
            DataField       =   "AGE PEriod 3"
            Caption         =   "Period 3"
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
            DataField       =   "AGE Period 4"
            Caption         =   "Period 4"
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
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1035.213
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1035.213
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1005.165
            EndProperty
         EndProperty
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Customer ID  "
         DataSource      =   "adoPrimaryRS"
         Height          =   255
         Index           =   50
         Left            =   480
         TabIndex        =   159
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Customer Name  "
         DataSource      =   "adoPrimaryRS"
         Height          =   255
         Index           =   49
         Left            =   480
         TabIndex        =   158
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Caption         =   "0 - 30"
         DataSource      =   "adoPrimaryRS"
         Height          =   255
         Index           =   48
         Left            =   480
         TabIndex        =   157
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Caption         =   "30 - 60"
         DataSource      =   "adoPrimaryRS"
         Height          =   255
         Index           =   47
         Left            =   2400
         TabIndex        =   156
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Caption         =   "60 - 90"
         DataSource      =   "adoPrimaryRS"
         Height          =   255
         Index           =   46
         Left            =   4320
         TabIndex        =   155
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Caption         =   "Over 90"
         DataSource      =   "adoPrimaryRS"
         Height          =   255
         Index           =   45
         Left            =   6240
         TabIndex        =   154
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Caption         =   "Total"
         DataSource      =   "adoPrimaryRS"
         Height          =   255
         Index           =   44
         Left            =   8160
         TabIndex        =   153
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Purchases  "
         DataSource      =   "adoPrimaryRS"
         Height          =   255
         Index           =   35
         Left            =   960
         TabIndex        =   152
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Purchases  "
         DataSource      =   "adoPrimaryRS"
         Height          =   255
         Index           =   36
         Left            =   960
         TabIndex        =   151
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Payments  "
         DataSource      =   "adoPrimaryRS"
         Height          =   255
         Index           =   37
         Left            =   960
         TabIndex        =   150
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Payments  "
         DataSource      =   "adoPrimaryRS"
         Height          =   255
         Index           =   38
         Left            =   960
         TabIndex        =   149
         Top             =   3720
         Width           =   1455
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Caption         =   "YTD"
         DataSource      =   "adoPrimaryRS"
         Height          =   255
         Index           =   40
         Left            =   2400
         TabIndex        =   148
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Caption         =   "Lifetime"
         DataSource      =   "adoPrimaryRS"
         Height          =   255
         Index           =   41
         Left            =   6240
         TabIndex        =   147
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Caption         =   "Last Year"
         DataSource      =   "adoPrimaryRS"
         Height          =   255
         Index           =   42
         Left            =   4320
         TabIndex        =   146
         Top             =   2400
         Width           =   1815
      End
   End
   Begin VB.Label lblTop 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Vendor/Remit  Data"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   54
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "frm_AP_Vendor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ADOprimaryrs As ADODB.Recordset
Attribute ADOprimaryrs.VB_VarHelpID = -1
Dim ADOTransRS As ADODB.Recordset
Dim ADOaging As ADODB.Recordset

Dim db As ADODB.Connection
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean

Dim sql As String
Dim TempStr As String
Dim WhichField As String
Dim WhichFields As String
Dim CurrVend As String

Dim VendorP1Balance As Currency
Dim VendorP2Balance As Currency
Dim VendorP3Balance As Currency
Dim VendorP4Balance As Currency
Dim VendorPtotalBalance As Currency

Public Sub CallByUserVendor(VendorID As String)
    Me.Show
    If mbAddNewFlag = False Then
        cmdAdd_Click
    End If
    txtFields(34) = VendorID
    txtFields(0) = VendorID
End Sub


Private Sub btcbRefresh_Click(Index As Integer)
    Dim tmp As String
    tmp = cbfields(Index).Text
    loadCombo Index
    cbfields(Index).Text = tmp
End Sub

Private Sub btVenID_Click()
    Dim ghead As String
    Dim fhead As String
 
    ghead = "Vendor"
    fhead = "ID//Name"
    AllLookup.ToWhichRecord ADOprimaryrs, ghead, fhead
    'AllLookup.Show vbModal
End Sub

Private Sub cbfields_LostFocus(Index As Integer)
Select Case Index
Case 1
   CheckCombo cbfields(Index), "[LIST VENDOR Types]", "[LIST Vendor Types]", db, True
Case 2
   CheckCombo cbfields(Index), "[LIST PAY Description]", "[LIST Payment Terms]", db, True
End Select
End Sub

Private Sub cdmShowAll_Click()
    Set grdDataGrid.DataSource = ADOprimaryrs
    cmdAdd.Enabled = False
    cmdUpdate.Enabled = False
    cmdDelete.Enabled = False
    cmdRefresh.Enabled = False
    lblTop.Caption = "Vendor Data"
    frShowAll.Visible = True
    frShowAll.ZOrder 0
End Sub

Private Sub chkFields_Click(Index As Integer)
If Index = 1 Then
    If chkFields(1).Value = 0 Then
        txtFields(1).Visible = False
        lblLabels(3).Visible = False
    Else
        txtFields(1).Visible = True
        lblLabels(3).Visible = True
    End If
    'GetTextColor Me
End If
End Sub

Private Sub cmdAgingDetails_Click()
  Set DataGrid1.DataSource = Nothing
    
  If ADOaging Is Nothing Then
  Else
    ADOaging.Close
    Set ADOaging = Nothing
  End If
  'AgingDetails txtfields(22).Text, db
  
  ShowStatus True
    Set ADOaging = New ADODB.Recordset
    ADOaging.Open "SELECT[AGE PO Doc Ext No],[AGE Start Date],[AGE Orig Amount]," & _
    "[AGE PEriod 1],[AGE Period 2],[AGE PEriod 3],[AGE Period 4] from " & _
    "[AGE Aging Purchase Work] WHERE [AGE Vendor ID]='" & txtFields(57).Text & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
  Set DataGrid1.DataSource = ADOaging
  DataGrid1.Visible = True
  
  ShowStatus False
End Sub

Private Sub cmdBack_Click()
If cmdBack.Caption = "Back" Then
    Set grdDataGrid.DataSource = Nothing
    cmdAdd.Enabled = True
    cmdUpdate.Enabled = True
    cmdDelete.Enabled = True
    cmdRefresh.Enabled = True
    lblTop.Caption = "Vendor/Remit  Data"
    frShowAll.Visible = False
Else
    Unload Me
End If
End Sub

Private Sub cmdBack1_Click()
    Set DataGrid2.DataSource = Nothing
    AP_VendorTrans.Visible = False
    If ADOTransRS Is Nothing Then
    Else
        ADOTransRS.Close
        Set ADOTransRS = Nothing
    End If
    cmdAdd.Enabled = True
    cmdUpdate.Enabled = True
    cmdDelete.Enabled = True
    cmdRefresh.Enabled = True
    lblTop.Caption = "Vendor/Remit  Data"
End Sub

Public Sub ShowListVendor()
    frm_AP_Vendor.Show
    cmdBack.Caption = "Close"
    cdmShowAll_Click
End Sub

Private Sub cmdDate_Click(Index As Integer)
Select Case Index
Case 0
    Menu_Calendar.WhoCallMe True, 1670
Case 1
    Menu_Calendar.WhoCallMe True, 1680
End Select
End Sub

Private Sub cmdGL_Click()
    Dim SQLstatement As String
    Dim No As Integer
    Dim ghead As String
    Dim fhead As String

    No = 38
    SQLstatement = "select [GL COA Account No], [GL COA Account Name]" & _
                    "from [GL Chart Of Accounts]"
    ghead = "Account Description"
    fhead = "Account No//Description"
    
    AllLookup.GetWhichTable No, SQLstatement, ghead, fhead, db
    'AllLookup.Show vbModal
    txtFields(30).SetFocus  ' trigger event adFirstChange

End Sub

Private Sub cmdSearch_Click()
If ADOTransRS Is Nothing Then
Else
    If ADOTransRS.RecordCount = 0 Then Exit Sub
    SearchRECORD ADOTransRS, DataGrid2, txtFields(35).Text, lblLabels(33).Caption, WhichField, "AP PO Ext Document No"
End If
End Sub

Private Sub cmdShow_Click()
If txtFields(32).Text = "" Or txtFields(33).Text = "" Then
    MsgBox "Please select start and the end date before executing the process"
    Exit Sub
End If
If ADOTransRS Is Nothing Then
Else
    ADOTransRS.Close
    Set ADOTransRS = Nothing
End If
Set ADOTransRS = New ADODB.Recordset
    
    ADOTransRS.Open TempStr & " AND [AP PO Date] BETWEEN #" & txtFields(32).Text & "# AND #" & txtFields(33).Text & "#", db, adOpenKeyset, adLockReadOnly, adCmdText
    
    Set DataGrid2.DataSource = ADOTransRS
    If ADOTransRS.RecordCount = 0 Then MsgBox "There is no transaction yet with " & txtFields(0)
End Sub

Private Sub Command1_Click()
Dim i As Integer

For i = 17 To 29
    txtFields(i).Text = txtFields(i - 15).Text
Next
End Sub

Private Sub Command2_Click()
If ADOTransRS Is Nothing Then
Else
    ADOTransRS.Close
    Set ADOTransRS = Nothing
End If
Set ADOTransRS = New ADODB.Recordset
    
    ADOTransRS.Open TempStr, db, adOpenKeyset, adLockReadOnly, adCmdText
    
    Set DataGrid2.DataSource = ADOTransRS
    If ADOTransRS.RecordCount = 0 Then MsgBox "There is no transaction yet with " & txtFields(0)
End Sub

Private Sub Command3_Click()
    cmdAdd.Enabled = True
    cmdUpdate.Enabled = True
    cmdDelete.Enabled = True
    cmdRefresh.Enabled = True
    lblTop.Caption = "Vendor/Remit  Data"
    AP_VENDORfinancials.Visible = False
End Sub

Private Sub Command4_Click()
If ADOprimaryrs Is Nothing Then
Else
    If ADOprimaryrs.RecordCount = 0 Then Exit Sub
    SearchRECORD ADOprimaryrs, grdDataGrid, txtFields(45).Text, lblLabels(34).Caption, WhichFields, "AP VEN ID"
End If
End Sub


Private Sub Command5_Click()

  'Load Vendor Balances

  'On Error Resume Next
  If DataGrid1.Visible = True Then
    Set DataGrid1.DataSource = Nothing
  End If
  ShowStatus True
  
  With ADOprimaryrs
  
  Call GetVendorFinancials(txtFields(57).Text, db, VendorP1Balance, VendorP2Balance, VendorP3Balance, VendorP4Balance, VendorPtotalBalance)
  ![AP VEN Financial Period 1] = VendorP1Balance
  ![AP VEN Financial Period 2] = VendorP2Balance
  ![AP VEN Financial Period 3] = VendorP3Balance
  ![AP VEN Financial Period 4] = VendorP4Balance
  ![AP VEN Financial Total] = VendorPtotalBalance

  'Get some invoice information
  'YTD
  
  'What is first day of year
  Dim DayOne As Date
  DayOne = FormatDate("1/01/" & Format(Now, "yyyy"))
  
  Dim Purchases#
  Dim Refunds#
  Purchases# = IIf(IsNull(SumRecord("[AP PO Total Amount]", "[AP Purchase]", db, "[AP PO Vendor ID] = '" & txtFields(57).Text & "' AND [AP PO Posted YN] = TRUE AND [AP PO Date] >= #" & DayOne & "# AND [AP PO Document Type] in ('Receiving','Voucher','Beginning Balance')")), 0, SumRecord("[AP PO Total Amount]", "[AP Purchase]", db, "[AP PO Vendor ID] = '" & txtFields(57).Text & "' AND [AP PO Posted YN] = TRUE AND [AP PO Date] >= #" & DayOne & "# AND [AP PO Document Type] in ('Receiving','Voucher','Beginning Balance')"))
  Refunds# = IIf(IsNull(SumRecord("[AP PO Total Amount]", "[AP Purchase]", db, "[AP PO Vendor ID] = '" & txtFields(57).Text & "' AND [AP PO Posted YN] = TRUE AND [AP PO Date] >= #" & DayOne & "# AND [AP PO Document Type] in ('Credit Memo')")), 0, SumRecord("[AP PO Total Amount]", "[AP Purchase]", db, "[AP PO Vendor ID] = '" & txtFields(57).Text & "' AND [AP PO Posted YN] = TRUE AND [AP PO Date] >= #" & DayOne & "# AND [AP PO Document Type] in ('Credit Memo')"))
  txtFields(47).Text = FormatCurr(Purchases# - Refunds#)
  txtFields(40).Text = FormatCurr(SumRecord("[AP PAY Amount]", "[AP Payment Header]", db, "[AP PAY Vendor No] = '" & txtFields(57).Text & "' AND [AP PAY Posted YN] = TRUE AND [AP PAY Transaction Date] >= #" & DayOne & "# AND [AP PAY Void] = FALSE AND [AP PAY Type] <> 'Credit Memo'"))
  txtFields(37).Text = CountRecord("[AP PO Total Amount]", "[AP Purchase]", db, "[AP PO Vendor ID] = '" & txtFields(57).Text & "' AND [AP PO Posted YN] = TRUE AND [AP PO Date] >= #" & DayOne & "# AND [AP PO Document Type] in ('Receiving','Voucher','Beginning Balance')")
  txtFields(43).Text = CountRecord("[AP PAY Amount]", "[AP Payment Header]", db, "[AP PAY Vendor No] = '" & txtFields(57).Text & "' AND [AP PAY Posted YN] = TRUE AND [AP PAY Transaction Date] >= #" & DayOne & "# AND [AP PAY Void] = FALSE AND [AP PAY Type] <> 'Credit Memo'")

  If txtFields(47).Text = "" Then txtFields(47).Text = "$0.00"
  If txtFields(40).Text = "" Then txtFields(40).Text = "$0.00"
  If txtFields(37).Text = "" Then txtFields(37).Text = 0
  If txtFields(43).Text = "" Then txtFields(43).Text = 0
  
  'Last Year
  Dim LastDay As Variant
  LastDay = DateAdd("d", -1, DayOne)
  DayOne = DateAdd("yyyy", -1, DayOne)
  Purchases# = IIf(IsNull(SumRecord("[AP PO Total Amount]", "[AP Purchase]", db, "[AP PO Vendor ID] = '" & txtFields(57).Text & "' AND [AP PO Posted YN] = TRUE AND [AP PO Date] BETWEEN #" & DayOne & "# AND #" & LastDay & "# AND [AP PO Document Type] in ('Receiving','Voucher','Beginning Balance')")), 0, SumRecord("[AP PO Total Amount]", "[AP Purchase]", db, "[AP PO Vendor ID] = '" & txtFields(57).Text & "' AND [AP PO Posted YN] = TRUE AND [AP PO Date] >= #" & DayOne & "# AND [AP PO Document Type] in ('Receiving','Voucher','Beginning Balance')"))
  Refunds# = IIf(IsNull(SumRecord("[AP PO Total Amount]", "[AP Purchase]", db, "[AP PO Vendor ID] = '" & txtFields(57).Text & "' AND [AP PO Posted YN] = TRUE AND [AP PO Date] BETWEEN #" & DayOne & "# AND #" & LastDay & "# AND [AP PO Document Type] in ('Credit Memo')")), 0, SumRecord("[AP PO Total Amount]", "[AP Purchase]", db, "[AP PO Vendor ID] = '" & txtFields(57).Text & "' AND [AP PO Posted YN] = TRUE AND [AP PO Date] >= #" & DayOne & "# AND [AP PO Document Type] in ('Credit Memo')"))
  ![AP VEN Purchase Last Year] = FormatCurr(Purchases# - Refunds#)
  ![AP VEN Payments Last Year] = FormatCurr(SumRecord("[AP PAY Amount]", "[AP Payment Header]", db, "[AP PAY Vendor No] = '" & txtFields(57).Text & "' AND [AP PAY Posted YN] = TRUE AND [AP PAY Transaction Date] BETWEEN #" & DayOne & "# AND #" & LastDay & "# AND [AP PAY Void] = FALSE AND [AP PAY Type] <> 'Credit Memo'"))
  ![AP VEN Purchase Number Last Year] = CountRecord("[AP PO Total Amount]", "[AP Purchase]", db, "[AP PO Vendor ID] = '" & txtFields(57).Text & "' AND [AP PO Posted YN] = TRUE AND [AP PO Date] BETWEEN #" & DayOne & "# AND #" & LastDay & "# AND [AP PO Document Type] in ('Receiving','Voucher','Beginning Balance')")
  ![AP VEN Payment Number Last Year] = CountRecord("[AP PAY Amount]", "[AP Payment Header]", db, "[AP PAY Vendor No] = '" & txtFields(57).Text & "' AND [AP PAY Posted YN] = TRUE AND [AP PAY Transaction Date] BETWEEN #" & DayOne & "# AND #" & LastDay & "# AND [AP PAY Void] = FALSE AND [AP PAY Type] <> 'Credit Memo'")
  
  If txtFields(49).Text Then txtFields(49).Text = "$0.00"
  If txtFields(38).Text Then txtFields(38).Text = "$0.00"
  If txtFields(46).Text Then txtFields(46).Text = 0
  If txtFields(41).Text Then txtFields(41).Text = 0

  'Lifetime
  Purchases# = IIf(IsNull(SumRecord("[AP PO Total Amount]", "[AP Purchase]", db, "[AP PO Vendor ID] = '" & txtFields(57).Text & "' AND [AP PO Posted YN] = TRUE AND [AP PO Document Type] in ('Receiving','Voucher','Beginning Balance')")), 0, SumRecord("[AP PO Total Amount]", "[AP Purchase]", db, "[AP PO Vendor ID] = '" & txtFields(57).Text & "' AND [AP PO Posted YN] = TRUE AND [AP PO Date] >= #" & DayOne & "# AND [AP PO Document Type] in ('Receiving','Voucher','Beginning Balance')"))
  Refunds# = IIf(IsNull(SumRecord("[AP PO Total Amount]", "[AP Purchase]", db, "[AP PO Vendor ID] = '" & txtFields(57).Text & "' AND [AP PO Posted YN] = TRUE AND [AP PO Document Type] in ('Credit Memo')")), 0, SumRecord("[AP PO Total Amount]", "[AP Purchase]", db, "[AP PO Vendor ID] = '" & txtFields(57).Text & "' AND [AP PO Posted YN] = TRUE AND [AP PO Date] >= #" & DayOne & "# AND [AP PO Document Type] in ('Credit Memo')"))
  ![AP VEN Purchase Lifetime] = FormatCurr(Purchases# - Refunds#)
  ![AP VEN Payments Lifetime] = FormatCurr(SumRecord("[AP PAY Amount]", "[AP Payment Header]", db, "[AP PAY Vendor No] = '" & txtFields(57).Text & "' AND [AP PAY Posted YN] = TRUE AND [AP PAY Void] = FALSE AND [AP PAY Type] <> 'Credit Memo'"))
  ![AP VEN Purchase Number Lifetime] = CountRecord("[AP PO Total Amount]", "[AP Purchase]", db, "[AP PO Vendor ID] = '" & txtFields(57).Text & "' AND [AP PO Posted YN] = TRUE AND [AP PO Document Type] in ('Receiving','Voucher','Beginning Balance')")
  ![AP VEN Payment Number Lifetime] = CountRecord("[AP PAY Amount]", "[AP Payment Header]", db, "[AP PAY Vendor No] = '" & txtFields(57).Text & "' AND [AP PAY Posted YN] = TRUE AND [AP PAY Void] = FALSE AND [AP PAY Type] <> 'Credit Memo'")
  
  If txtFields(38).Text Then txtFields(38).Text = "$0.00"
  If txtFields(39).Text Then txtFields(39).Text = "$0.00"
  If txtFields(36).Text Then txtFields(36).Text = 0
  If txtFields(42).Text Then txtFields(42).Text = 0
  
  .Update
  End With
  ShowStatus False
  If DataGrid1.Visible = True Then cmdAgingDetails_Click
End Sub

Private Sub DataGrid2_HeadClick(ByVal ColIndex As Integer)
If ADOTransRS.RecordCount = 0 Then Exit Sub
    lblLabels(33) = DataGrid2.Columns(ColIndex).Caption
    WhichField = DataGrid2.Columns(ColIndex).DataField
    ADOTransRS.Close
    Set ADOTransRS = Nothing
    Set ADOTransRS = New ADODB.Recordset
    ADOTransRS.Open TempStr & " ORDER BY [" & WhichField & "]", db, adOpenKeyset, adLockOptimistic, adCmdText
    Set DataGrid2.DataSource = ADOTransRS
End Sub

Private Sub FinBut_Click()
    cmdAdd.Enabled = False
    cmdUpdate.Enabled = False
    cmdDelete.Enabled = False
    cmdRefresh.Enabled = False
    lblTop.Caption = "Vendor Financial Data"
    AP_VENDORfinancials.Visible = True
    AP_VENDORfinancials.ZOrder 0
End Sub

Private Sub Form_Load()
mbAddNewFlag = False

ShowStatus True
'On Error GoTo FormErr
    Set db = New ADODB.Connection
    db.CursorLocation = adUseClient
    db.Open gblADOProvider
  
  'Dim sql As String
  sql = "select [AP VEN ID],[AP VEN Name],[AP VEN Address 1],[AP VEN Address 2]," & _
    "[AP VEN City],[AP VEN State],[AP VEN Credit Limit],[AP VEN Postal],[AP VEN Country],[AP VEN Contact]," & _
    "[AP VEN Phone],[AP VEN Phone Ext],[AP VEN Fax],[AP VEN Department]," & _
    "[AP VEN Default Department],[AP VEN Type],[AP VEN Tax ID],[AP VEN Payment Hold]," & _
    "[AP VEN Inactive YN],[AP VEN Remit Name],[AP VEN Remit Address 1]," & _
    "[AP VEN Remit Address 2],[AP VEN Remit City],[AP VEN Remit State]," & _
    "[AP VEN Remit Postal],[AP VEN Remit Country],[AP VEN Remit Contact]," & _
    "[AP VEN Remit Phone],[AP VEN Remit Phone Ext],[AP VEN Remit Fax]," & _
    "[AP VEN Financial Period 1],[AP VEN Financial Period 2],[AP VEN Financial Period 3]," & _
    "[AP VEN Financial Period 4],[AP VEN Financial Total],[AP VEN Payments YTD]," & _
    "[AP VEN Purchase YTD],[AP VEN Purchase Number YTD],[AP VEN Payment Number YTD]," & _
    "[AP VEN Purchase Last Year],[AP VEN Payments Last Year],[AP VEN Purchase Number Last Year],[AP VEN Payment Number Last Year]," & _
    "[AP VEN Purchase Lifetime],[AP VEN Payments Lifetime],[AP VEN Purchase Number Lifetime],[AP VEN Payment Number Lifetime]," & _
    "[AP VEN Default GL],[AP VEN Default Terms],[AP VEN 1099],[AP VEN Notes],[AP VEN Custom 1]," & _
    "[AP VEN Custom 2],[AP VEN Custom 3],[AP VEN Custom 4] from [AP Vendor]"
    
  Set ADOprimaryrs = New ADODB.Recordset
  ADOprimaryrs.Open sql, db, adOpenStatic, adLockOptimistic
  
  Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtFields
      Set oText.DataSource = ADOprimaryrs
    If oText.DataField <> "" Then
        If ADOprimaryrs("" & oText.DataField & "").Type = 202 Then oText.MaxLength = ADOprimaryrs("" & oText.DataField & "").DefinedSize
    End If
  Next

  Dim oCheck As CheckBox
    'Bind the Check boxes to the data provider
  For Each oCheck In Me.chkFields
    Set oCheck.DataSource = ADOprimaryrs
  Next

  Dim oCombo As ComboBox
  'Bind datacombos to the data provider
  For Each oCombo In Me.cbfields
    Set oCombo.DataSource = ADOprimaryrs
  Next
  
  Me.Width = 10755
  Me.Height = 8775
  loadCombo
  
  If CheckNewDB(ADOprimaryrs, "Vendor") = True Then
    cmdAdd_Click
  End If
  
  GetTextColor Me
  mbDataChanged = False
ShowStatus False
Exit Sub
FormErr:
  MsgBox Err.Description
  ShowStatus False
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
  
  Me.Width = 10755
  Me.Height = 8775
  
SkipResize:
  frPrimary.Left = (Me.ScaleWidth - frPrimary.Width) / 2
  lblTop.Left = frPrimary.Left
  lblTop.Width = frPrimary.Width
  frPrimary.Top = (Me.ScaleHeight - frPrimary.Height - picButtons.Height - picStatBox.Height) / 2 + 230
  
  frShowAll.Left = frPrimary.Left
  frShowAll.Top = frPrimary.Top
  AP_VendorTrans.Left = frPrimary.Left
  AP_VendorTrans.Top = frPrimary.Top '
  AP_VENDORfinancials.Left = frPrimary.Left
  AP_VENDORfinancials.Top = frPrimary.Top '
  
  lblStatus.Width = Me.Width - 1500
  cmdNext.Left = lblStatus.Width + 700
  cmdLast.Left = cmdNext.Left + 340
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
    'updates the checklist Vendors
  ShowStatus True
    If ADOaging Is Nothing Then
    Else
      ADOaging.Close
      Set ADOaging = Nothing
    End If
    
      EndLoad db, ADOprimaryrs, "Vendors"
      If ADOprimaryrs.RecordCount > 0 Then
        If ADOprimaryrs.EditMode <> 0 Then
          ADOprimaryrs.CancelUpdate
        End If
      End If
      ADOprimaryrs.Close
      Set ADOprimaryrs = Nothing
      db.Close
      Set db = Nothing
      Set frm_AP_Vendor = Nothing
  ShowStatus False
  Exit Sub
FormErr:
  MsgBox Err.Description
  ShowStatus False
End Sub

Private Sub adoPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This will display the current record position for this recordset
  lblStatus.Caption = "Record: " & CStr(ADOprimaryrs.AbsolutePosition) & " of " & CStr(ADOprimaryrs.RecordCount)
  If ADOprimaryrs.BOF Or ADOprimaryrs.EOF Then Exit Sub
  CurrVend = ADOprimaryrs![AP VEN ID] & ""
  txtFields(34).Text = CurrVend
  If IsNull(ADOprimaryrs![AP VEN Credit Limit]) Then
    ADOprimaryrs![AP VEN Credit Limit] = "$0.00"
    If mbAddNewFlag = False Then ADOprimaryrs.Update
  End If
  chkFields_Click 1
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
On Error GoTo AddErr
  With ADOprimaryrs
    If cmdAdd.Caption = "&Add" Then
        If .RecordCount > 0 And Not mbAddNewFlag Then
            mvBookMark = .Bookmark
        End If
        mbAddNewFlag = True
        .AddNew
        lblStatus.Caption = "Add record"
        txtFields(0).Enabled = True
        cmdAdd.Caption = "&Cancel"
    Else
        mbAddNewFlag = False
        .CancelUpdate
        txtFields(0).Enabled = False
        cmdAdd.Caption = "&Add"
        If .RecordCount > 0 Then
            If mvBookMark > 0 Then
                .Bookmark = mvBookMark
            Else
                .MoveLast
            End If
        End If
    End If
    
    'set to controls appropriately
    btVenID.Enabled = Not mbAddNewFlag
    cmdDelete.Enabled = Not mbAddNewFlag
    cmdRefresh.Enabled = Not mbAddNewFlag
  End With
  GetTextColor Me
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
On Error GoTo DeleteErr
  With ADOprimaryrs
    If .RecordCount = 0 Then Exit Sub   ' no records maa....
    If .EditMode = False Then
        .Delete
        .MoveNext
        If .RecordCount = 0 Then  ' no more records
            cmdUpdate.Enabled = False
            cmdDelete.Enabled = False
            cmdRefresh.Enabled = False
            .Requery
            Exit Sub
        ElseIf .EOF Then
            .MoveLast
        End If
        If Not (.BOF Or .EOF) Then mvBookMark = .Bookmark
    Else
        MsgBox "Must update or refresh record before deleting.", vbCritical, _
            "Delete Error."
    End If
  End With
  Exit Sub
DeleteErr:
  MsgBox Err.Description
End Sub

Private Sub cmdRefresh_Click()
'This is only needed for multi user apps
On Error GoTo RefreshErr

If ADOprimaryrs.RecordCount > 0 Then
  
  If ADOprimaryrs.RecordCount > 1 Then
        If ADOprimaryrs.EditMode <> 0 Then
            ADOprimaryrs.UpdateBatch adAffectAll
        End If
        
        mvBookMark = ADOprimaryrs.Bookmark
          ADOprimaryrs.Requery
        ADOprimaryrs.Bookmark = mvBookMark
  Else
        ADOprimaryrs.UpdateBatch adAffectAll
        ADOprimaryrs.Requery
  End If

End If
  Exit Sub
RefreshErr:
  MsgBox Err.Description
  Resume Next
End Sub

Private Sub cmdUpdate_Click()
    On Error GoTo UpdateErr

    With ADOprimaryrs
        If .RecordCount = 0 Then Exit Sub 'no records to update
        If Trim(txtFields(0).Text) <> "" Then
        Dim oTxt As Control
          For Each oTxt In Me.Controls
          If TypeOf Ctrl Is TextBox Then
            If oTxt.Text = "" Then
              If ADOprimaryrs("" & oTxt.DataField & "").Type = 203 Or ADOprimaryrs("" & oTxt.DataField & "").Type = 202 Then oTxt.Text = " "
            End If
          End If
          Next
        End If
        .Update
        If mbAddNewFlag Then 'requery to get default value assigned by database
            .Requery
            .MoveLast
            mbAddNewFlag = False
        End If
        
        'reenable the necessary buttons
        cmdAdd.Caption = "&Add"
        txtFields(0).Enabled = False
        cmdDelete.Enabled = True
        cmdRefresh.Enabled = True
        btVenID.Enabled = True
    End With

  mbEditFlag = False
  GetTextColor Me
  mbDataChanged = False

  Exit Sub
UpdateErr:
  MsgBox Err.Description
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
  cmdDelete.Visible = bVal
  cmdClose.Visible = bVal
  cmdRefresh.Visible = bVal
  cmdNext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdLast.Enabled = bVal
  cmdPrevious.Enabled = bVal
End Sub

Private Sub loadCombo(Optional Index As Integer)
    Index = IIf(Index > 0, Index, 0)
    Select Case Index
    Case 0
        ComboInit cbfields(1), lblfields(1), "select [LIST VENDOR Types] from [LIST Vendor Types]"
        ComboInit cbfields(2), lblfields(2), "select [LIST PAY Description] from [LIST Payment Terms]"
    Case 1
        ComboInit cbfields(1), lblfields(1), "select [LIST VENDOR Types] from [LIST Vendor Types]"
    Case 2
        ComboInit cbfields(2), lblfields(2), "select [LIST PAY Description] from [LIST Payment Terms]"
    End Select
    
End Sub

Private Sub grdDataGrid_HeadClick(ByVal ColIndex As Integer)
If ADOprimaryrs.RecordCount = 0 Then Exit Sub
    lblLabels(34) = grdDataGrid.Columns(ColIndex).Caption
    WhichFields = grdDataGrid.Columns(ColIndex).DataField
    ADOprimaryrs.Close
    Set ADOprimaryrs = Nothing
    Set ADOprimaryrs = New ADODB.Recordset
        
    ADOprimaryrs.Open sql & " Order by [" & WhichFields & "]", db, adOpenStatic, adLockOptimistic, adCmdText
    Set grdDataGrid.DataSource = ADOprimaryrs
End Sub

Private Sub Tranbut_Click()
Set ADOTransRS = New ADODB.Recordset

TempStr = "SELECT [AP PO Posted YN],[AP PO Date],[AP PO Document Type]," & _
"[AP PO Vendor Invoice No],[AP PO Ext Document No],[AP PO Payment Terms]," & _
"[AP PO Amount Paid],[AP PO Total Amount] FROM [AP Purchase] " & _
"WHERE [AP PO Vendor ID]='" & txtFields(34).Text & "'"

ADOTransRS.Open TempStr, db, adOpenKeyset, adLockOptimistic, adCmdText
Set DataGrid2.DataSource = ADOTransRS
AP_VendorTrans.Visible = True
AP_VendorTrans.ZOrder 0

    cmdAdd.Enabled = False
    cmdUpdate.Enabled = False
    cmdDelete.Enabled = False
    cmdRefresh.Enabled = False
    lblTop.Caption = "Vendor Transaction"

End Sub

Private Sub txtfields_GotFocus(Index As Integer)
    TxtGotFocus txtFields(Index)
End Sub

Private Sub txtfields_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 16 Then
     keyResponse = CtrlValidate(KeyAscii, "0123456789.")
     If keyResponse = True Then
     Else
        KeyAscii = 0
     End If
End If
End Sub

Private Sub txtFields_LostFocus(Index As Integer)
Select Case Index
Case 34
    Dim txtVendID As String
    txtVendID = txtFields(34)
    If txtFields(34) = "" And mbAddNewFlag = False Then
        txtFields(34) = CurrVend
    ElseIf txtFields(34) <> "" And mbAddNewFlag = True Then
        If CheckDocument("SELECT [AP VEN ID] FROM [AP VENDOR] WHERE [AP VEN ID]='" & txtVendID & "'", db, False) = False Then
            MsgBox txtVendID & " is already exist", vbInformation, "Information"
            txtFields(34) = ""
            txtFields(0) = txtFields(34)
            Exit Sub
        Else
            txtFields(0) = txtVendID
        End If
    End If
    
    If txtFields(34) = CurrVend Then Exit Sub
    
    With ADOprimaryrs
      If .RecordCount > 0 And mbAddNewFlag = False Then
        mvBookMark = .Bookmark
        .MoveFirst
        .Find "[AP VEN ID]='" & txtVendID & "'"
        If .EOF Then
            Dim Response As Integer
            Response = MsgBox(txtVendID & " is a new input. Would you like to add it into the database", vbYesNo, "Information")
            If Response = vbYes Then
                mbAddNewFlag = True
                cmdAdd_Click
                txtFields(34) = txtVendID
            Else
                .Bookmark = mvBookMark
                txtFields(34) = txtFields(0)
            End If
            txtFields(34).SetFocus
        End If
      End If
    'Else
    End With
End Select
End Sub
