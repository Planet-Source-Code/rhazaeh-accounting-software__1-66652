VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_AR_Customer 
   Caption         =   "Customer Data"
   ClientHeight    =   7095
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   10365
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7095
   ScaleWidth      =   10365
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   660
      Left            =   0
      ScaleHeight     =   660
      ScaleWidth      =   10365
      TabIndex        =   65
      Top             =   6435
      Width           =   10365
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Picture         =   "frm_AR_Customer.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   92
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Picture         =   "frm_AR_Customer.frx":0342
         Style           =   1  'Graphical
         TabIndex        =   91
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Picture         =   "frm_AR_Customer.frx":0684
         Style           =   1  'Graphical
         TabIndex        =   90
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Picture         =   "frm_AR_Customer.frx":09C6
         Style           =   1  'Graphical
         TabIndex        =   89
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   300
         Left            =   4320
         TabIndex        =   41
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   300
         Left            =   3240
         TabIndex        =   40
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   300
         Left            =   2160
         TabIndex        =   39
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   300
         Left            =   1080
         TabIndex        =   38
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   300
         Left            =   0
         TabIndex        =   37
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   93
         Top             =   360
         Width           =   3360
      End
   End
   Begin VB.Frame frPrimary 
      Height          =   5895
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   10335
      Begin VB.TextBox txtfields 
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   34
         Left            =   1320
         TabIndex        =   131
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtfields 
         DataField       =   "AR EMail Address"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   21
         Left            =   1320
         TabIndex        =   12
         Top             =   3360
         Width           =   3375
      End
      Begin VB.TextBox txtfields 
         DataField       =   "AR CUST Web Page"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   19
         Left            =   1320
         TabIndex        =   4
         Top             =   1200
         Width           =   3375
      End
      Begin VB.CommandButton cdmShowAll 
         Caption         =   "List All"
         Height          =   795
         Left            =   9120
         Picture         =   "frm_AR_Customer.frx":0D08
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   4920
         Width           =   855
      End
      Begin VB.CommandButton ShipToBut 
         Caption         =   "Ship To"
         Height          =   795
         Left            =   8040
         Picture         =   "frm_AR_Customer.frx":1012
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   3840
         Width           =   855
      End
      Begin VB.CommandButton FinBut 
         Caption         =   "Financials"
         Height          =   795
         Left            =   8040
         Picture         =   "frm_AR_Customer.frx":131C
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   4920
         Width           =   855
      End
      Begin VB.TextBox txtfields 
         DataField       =   "AR CUST Discount %"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   20
         Left            =   9120
         TabIndex        =   30
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox txtfields 
         DataField       =   "AR CUST Tax ID No"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   12
         Left            =   6480
         TabIndex        =   17
         Top             =   480
         Width           =   1575
      End
      Begin VB.CheckBox chkFields 
         Alignment       =   1  'Right Justify
         Caption         =   "Statements (Y/N)"
         DataField       =   "AR CUST Statements YN"
         DataSource      =   "adoPrimaryRS"
         Height          =   255
         Index           =   21
         Left            =   6000
         TabIndex        =   32
         Top             =   4080
         Width           =   1575
      End
      Begin VB.TextBox txtfields 
         DataField       =   "AR CUST State"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   5
         Left            =   2880
         TabIndex        =   8
         Top             =   2280
         Width           =   615
      End
      Begin VB.TextBox txtfields 
         DataField       =   "AR CUST Sales Account"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   18
         Left            =   5640
         TabIndex        =   26
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox txtfields 
         DataField       =   "AR CUST Postal"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   6
         Left            =   3960
         TabIndex        =   9
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txtfields 
         DataField       =   "AR CUST Phone Ext"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   10
         Left            =   3480
         TabIndex        =   14
         Top             =   3720
         Width           =   1215
      End
      Begin VB.TextBox txtfields 
         DataField       =   "AR CUST Phone"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   9
         Left            =   1320
         TabIndex        =   13
         Top             =   3720
         Width           =   1575
      End
      Begin VB.TextBox txtfields 
         DataField       =   "AR CUST Notes"
         DataSource      =   "adoPrimaryRS"
         Height          =   1095
         Index           =   23
         Left            =   1320
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Top             =   4560
         Width           =   6255
      End
      Begin VB.TextBox txtfields 
         DataField       =   "AR CUST Name"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   1
         Left            =   1320
         TabIndex        =   3
         Top             =   840
         Width           =   3375
      End
      Begin VB.CheckBox chkFields 
         Alignment       =   1  'Right Justify
         Caption         =   "Finance Charge (Y/N)"
         DataField       =   "AR CUST Finance Charge YN"
         DataSource      =   "adoPrimaryRS"
         Height          =   255
         Index           =   22
         Left            =   5640
         TabIndex        =   31
         Top             =   3720
         Width           =   1935
      End
      Begin VB.TextBox txtfields 
         DataField       =   "AR CUST Fax"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   11
         Left            =   1320
         TabIndex        =   15
         Top             =   4080
         Width           =   1575
      End
      Begin VB.TextBox txtfields 
         DataField       =   "AR CUST Customer ID"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   0
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtfields 
         DataField       =   "AR CUST Credit Limit"
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
         Index           =   14
         Left            =   6480
         TabIndex        =   19
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox txtfields 
         DataField       =   "AR CUST Country"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   7
         Left            =   1320
         TabIndex        =   10
         Top             =   2640
         Width           =   1575
      End
      Begin VB.TextBox txtfields 
         DataField       =   "AR CUST Contact"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   8
         Left            =   1320
         TabIndex        =   11
         Top             =   3000
         Width           =   3375
      End
      Begin VB.TextBox txtfields 
         DataField       =   "AR CUST City"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   4
         Left            =   1320
         TabIndex        =   7
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox txtfields 
         DataField       =   "AR CUST Average Days To Pay"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   13
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txtfields 
         DataField       =   "AR CUST Address 2"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   3
         Left            =   1320
         TabIndex        =   6
         Top             =   1920
         Width           =   3375
      End
      Begin VB.TextBox txtfields 
         DataField       =   "AR CUST Address 1"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   2
         Left            =   1320
         TabIndex        =   5
         Top             =   1560
         Width           =   3375
      End
      Begin VB.CommandButton btCustID 
         Height          =   285
         Left            =   2880
         Picture         =   "frm_AR_Customer.frx":1626
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton btCustSalesAcc 
         Height          =   285
         Left            =   7200
         Picture         =   "frm_AR_Customer.frx":1930
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   1920
         Width           =   375
      End
      Begin VB.ComboBox cbfields 
         DataField       =   "AR CUST Payment Terms"
         DataSource      =   "adoprimaryrs"
         Height          =   315
         Index           =   1
         Left            =   5640
         TabIndex        =   20
         Top             =   2640
         Width           =   1575
      End
      Begin VB.ComboBox cbfields 
         DataField       =   "AR CUST Type"
         Height          =   315
         Index           =   2
         Left            =   8040
         TabIndex        =   22
         Top             =   2640
         Width           =   1575
      End
      Begin VB.ComboBox cbfields 
         DataField       =   "AR CUST Tax Group"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         DataSource      =   "adoprimaryrs"
         Height          =   315
         Index           =   3
         Left            =   5640
         TabIndex        =   24
         Top             =   3240
         Width           =   1575
      End
      Begin VB.ComboBox cbfields 
         DataField       =   "AR CUST SalesPerson"
         Height          =   315
         Index           =   4
         Left            =   8040
         TabIndex        =   28
         Top             =   3240
         Width           =   1575
      End
      Begin VB.CommandButton btcbRefresh 
         Height          =   285
         Index           =   1
         Left            =   7200
         Picture         =   "frm_AR_Customer.frx":1C3A
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Update the Ship Via"
         Top             =   2640
         Width           =   375
      End
      Begin VB.CommandButton btcbRefresh 
         Height          =   285
         Index           =   4
         Left            =   9600
         Picture         =   "frm_AR_Customer.frx":1F44
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Update the Ship Via"
         Top             =   3240
         Width           =   375
      End
      Begin VB.CommandButton btcbRefresh 
         Height          =   285
         Index           =   2
         Left            =   9600
         Picture         =   "frm_AR_Customer.frx":224E
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Update the Ship Via"
         Top             =   2640
         Width           =   375
      End
      Begin VB.CommandButton btcbRefresh 
         Height          =   285
         Index           =   3
         Left            =   7200
         Picture         =   "frm_AR_Customer.frx":2558
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Update the Ship Via"
         Top             =   3240
         Width           =   375
      End
      Begin VB.CommandButton Tranbut 
         Caption         =   "Sales"
         Height          =   795
         Left            =   9120
         Picture         =   "frm_AR_Customer.frx":2862
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   3840
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "E-Mail:"
         Height          =   255
         Index           =   20
         Left            =   240
         TabIndex        =   85
         Top             =   3360
         Width           =   975
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Web Page:"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   84
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label lblfields 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tax Group"
         DataSource      =   "adoPrimaryRS"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   5640
         TabIndex        =   83
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Discount (%):"
         DataSource      =   "adoPrimaryRS"
         Height          =   255
         Index           =   23
         Left            =   8040
         TabIndex        =   61
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label lblfields 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Customer Type"
         DataSource      =   "adoPrimaryRS"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   8040
         TabIndex        =   60
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Tax ID Number:"
         DataSource      =   "adoPrimaryRS"
         Height          =   255
         Index           =   21
         Left            =   5040
         TabIndex        =   59
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "State:"
         Height          =   255
         Index           =   18
         Left            =   2280
         TabIndex        =   58
         Top             =   2280
         Width           =   495
      End
      Begin VB.Label lblfields 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Salesman"
         DataSource      =   "adoPrimaryRS"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   8040
         TabIndex        =   57
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label lblfields 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "GL Sales Account"
         DataSource      =   "adoPrimaryRS"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   18
         Left            =   5640
         TabIndex        =   56
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Zip:"
         Height          =   255
         Index           =   15
         Left            =   3480
         TabIndex        =   55
         Top             =   2280
         Width           =   375
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Ext:"
         DataSource      =   "adoPrimaryRS"
         Height          =   255
         Index           =   14
         Left            =   3000
         TabIndex        =   54
         Top             =   3720
         Width           =   375
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Phone No:"
         Height          =   255
         Index           =   13
         Left            =   240
         TabIndex        =   53
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label lblfields 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Payment Terms"
         DataSource      =   "adoPrimaryRS"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   5640
         TabIndex        =   52
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Notes:"
         Height          =   255
         Index           =   11
         Left            =   240
         TabIndex        =   51
         Top             =   4560
         Width           =   975
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Name:"
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   50
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Fax Number:"
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   49
         Top             =   4080
         Width           =   975
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Customer ID:"
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   48
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Credit Limit:"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         DataSource      =   "adoPrimaryRS"
         Height          =   255
         Index           =   6
         Left            =   5040
         TabIndex        =   47
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Country:"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   46
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Contact:"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   45
         Top             =   3000
         Width           =   975
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "City:"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   44
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Average Pay Days"
         DataSource      =   "adoPrimaryRS"
         Height          =   255
         Index           =   2
         Left            =   5040
         TabIndex        =   43
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Address:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   42
         Top             =   1560
         Width           =   975
      End
   End
   Begin VB.Frame AR_Cust_Drill 
      Height          =   5895
      Left            =   0
      TabIndex        =   68
      Top             =   480
      Visible         =   0   'False
      Width           =   10335
      Begin VB.CommandButton cmdBack1 
         Caption         =   "&Back"
         Height          =   855
         Left            =   9120
         Picture         =   "frm_AR_Customer.frx":2B6C
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   600
         Width           =   975
      End
      Begin VB.Frame Frame1 
         Height          =   1335
         Left            =   120
         TabIndex        =   73
         Top             =   240
         Width           =   5055
         Begin VB.CommandButton Command1 
            Caption         =   "A&ll"
            Height          =   855
            Left            =   3960
            Picture         =   "frm_AR_Customer.frx":2E76
            Style           =   1  'Graphical
            TabIndex        =   82
            Top             =   360
            Width           =   975
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
            Index           =   17
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   78
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton cmdDate 
            Height          =   285
            Index           =   0
            Left            =   2400
            Picture         =   "frm_AR_Customer.frx":3180
            Style           =   1  'Graphical
            TabIndex        =   77
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
            Index           =   16
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   76
            Top             =   840
            Width           =   1455
         End
         Begin VB.CommandButton cmdDate 
            Height          =   285
            Index           =   1
            Left            =   2400
            Picture         =   "frm_AR_Customer.frx":348A
            Style           =   1  'Graphical
            TabIndex        =   75
            Top             =   840
            Width           =   375
         End
         Begin VB.CommandButton cmdShow 
            Caption         =   "&Execute"
            Height          =   855
            Left            =   2880
            Picture         =   "frm_AR_Customer.frx":3794
            Style           =   1  'Graphical
            TabIndex        =   74
            Top             =   360
            Width           =   975
         End
         Begin VB.Label lblLabels 
            Caption         =   "Start Date:"
            Height          =   255
            Index           =   27
            Left            =   120
            TabIndex        =   80
            Top             =   360
            Width           =   855
         End
         Begin VB.Label lblLabels 
            Caption         =   "End Date:"
            Height          =   255
            Index           =   26
            Left            =   120
            TabIndex        =   79
            Top             =   840
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1335
         Left            =   5280
         TabIndex        =   69
         Top             =   240
         Width           =   2535
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
            Index           =   15
            Left            =   240
            TabIndex        =   71
            Top             =   720
            Width           =   1335
         End
         Begin VB.CommandButton cmdSearch 
            Height          =   540
            Left            =   1680
            Picture         =   "frm_AR_Customer.frx":3A9E
            Style           =   1  'Graphical
            TabIndex        =   70
            Top             =   480
            Width           =   615
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Doc. Type"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   9
            Left            =   240
            TabIndex        =   72
            Top             =   480
            Width           =   1335
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   4095
         Left            =   120
         TabIndex        =   141
         Top             =   1680
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   7223
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
         Caption         =   "Customer Transaction"
         ColumnCount     =   8
         BeginProperty Column00 
            DataField       =   "AR SALE Posted YN"
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
            DataField       =   "AR SALE Date"
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
            DataField       =   "AR SALE Document Type"
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
            DataField       =   "AR SALE PO ID"
            Caption         =   "Cust. PO"
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
         BeginProperty Column05 
            DataField       =   "AR SALE Payment Terms"
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
            DataField       =   "AR SALE Amount Paid"
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
            DataField       =   "AR SALE Total"
            Caption         =   "Sales Total"
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
               ColumnWidth     =   1349.858
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1289.764
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1140.095
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame frShowAll 
      Height          =   5895
      Left            =   0
      TabIndex        =   63
      Top             =   480
      Visible         =   0   'False
      Width           =   10335
      Begin VB.Frame Frame2 
         Height          =   975
         Left            =   5400
         TabIndex        =   132
         Top             =   120
         Width           =   2535
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
            TabIndex        =   134
            Top             =   480
            Width           =   1335
         End
         Begin VB.CommandButton Command4 
            Height          =   540
            Left            =   1680
            Picture         =   "frm_AR_Customer.frx":3DA8
            Style           =   1  'Graphical
            TabIndex        =   133
            Top             =   240
            Width           =   615
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
            Index           =   12
            Left            =   240
            TabIndex        =   135
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         Height          =   795
         Left            =   9240
         Picture         =   "frm_AR_Customer.frx":40B2
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "Back"
         Height          =   795
         Left            =   8040
         Picture         =   "frm_AR_Customer.frx":43BC
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   240
         Width           =   975
      End
      Begin MSDataGridLib.DataGrid grdDataGrid 
         Bindings        =   "frm_AR_Customer.frx":46C6
         Height          =   4575
         Left            =   120
         TabIndex        =   66
         Top             =   1200
         Width           =   10065
         _ExtentX        =   17754
         _ExtentY        =   8070
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
         Caption         =   "Customer Data"
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "AR CUST Customer ID"
            Caption         =   "ID"
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
            DataField       =   "AR CUST Name"
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
            DataField       =   "AR CUST Contact"
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
            DataField       =   "AR CUST Phone"
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
            DataField       =   "AR CUST Phone Ext"
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
            DataField       =   "AR CUST Sales YTD"
            Caption         =   "Sales YTD"
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
            DataField       =   "AR CUST Notes"
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
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1904.882
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame AR_Cust_Financials 
      Height          =   5895
      Left            =   0
      TabIndex        =   86
      Top             =   480
      Visible         =   0   'False
      Width           =   10335
      Begin VB.CommandButton cmdAgingDetails 
         Caption         =   "Aging Detail"
         Height          =   375
         Left            =   120
         TabIndex        =   139
         Top             =   4560
         Width           =   1335
      End
      Begin VB.CommandButton cmdAveDays 
         Height          =   285
         Left            =   9600
         Picture         =   "frm_AR_Customer.frx":46E1
         Style           =   1  'Graphical
         TabIndex        =   138
         Top             =   960
         Width           =   375
      End
      Begin VB.CommandButton CmdPrtState 
         Caption         =   "Print Statement"
         Height          =   375
         Left            =   120
         TabIndex        =   136
         Top             =   4200
         Width           =   1335
      End
      Begin VB.CommandButton cdmAgeReceivable 
         Caption         =   "Age Receivable"
         Height          =   375
         Left            =   120
         TabIndex        =   137
         Top             =   3840
         Width           =   1335
      End
      Begin VB.TextBox txtfields 
         Alignment       =   1  'Right Justify
         DataField       =   "AR CUST Average Days To Pay"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   44
         Left            =   8280
         Locked          =   -1  'True
         TabIndex        =   114
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtfields 
         Alignment       =   1  'Right Justify
         DataField       =   "AR CUST Write Offs YTD"
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
         Index           =   43
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   113
         Top             =   3720
         Width           =   1815
      End
      Begin VB.TextBox txtfields 
         Alignment       =   1  'Right Justify
         DataField       =   "AR CUST Write Offs Lifetime"
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
         Index           =   42
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   112
         Top             =   3720
         Width           =   1815
      End
      Begin VB.TextBox txtfields 
         Alignment       =   1  'Right Justify
         DataField       =   "AR CUST Write Offs Last Year"
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
         Index           =   41
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   111
         Top             =   3720
         Width           =   1815
      End
      Begin VB.TextBox txtfields 
         Alignment       =   1  'Right Justify
         DataField       =   "AR CUST Sales YTD"
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
         TabIndex        =   110
         Top             =   3000
         Width           =   1815
      End
      Begin VB.TextBox txtfields 
         Alignment       =   1  'Right Justify
         DataField       =   "AR CUST Sales Lifetime"
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
         TabIndex        =   109
         Top             =   3000
         Width           =   1815
      End
      Begin VB.TextBox txtfields 
         Alignment       =   1  'Right Justify
         DataField       =   "AR CUST Sales Last Year"
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
         TabIndex        =   108
         Top             =   3000
         Width           =   1815
      End
      Begin VB.TextBox txtfields 
         Alignment       =   1  'Right Justify
         DataField       =   "AR CUST Payments YTD"
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
         Index           =   37
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   107
         Top             =   3360
         Width           =   1815
      End
      Begin VB.TextBox txtfields 
         Alignment       =   1  'Right Justify
         DataField       =   "AR CUST Payments Lifetime"
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
         Index           =   36
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   106
         Top             =   3360
         Width           =   1815
      End
      Begin VB.TextBox txtfields 
         Alignment       =   1  'Right Justify
         DataField       =   "AR CUST Payments Last Year"
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
         Index           =   35
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   105
         Top             =   3360
         Width           =   1815
      End
      Begin VB.TextBox txtfields 
         Alignment       =   1  'Right Justify
         DataField       =   "AR CUST Invoices YTD"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0;(0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   33
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   104
         Top             =   2640
         Width           =   1815
      End
      Begin VB.TextBox txtfields 
         Alignment       =   1  'Right Justify
         DataField       =   "AR CUST Invoices Lifetime"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0;(0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   32
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   103
         Top             =   2640
         Width           =   1815
      End
      Begin VB.TextBox txtfields 
         Alignment       =   1  'Right Justify
         DataField       =   "AR CUST Invoices Last Year"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0;(0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   31
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   102
         Top             =   2640
         Width           =   1815
      End
      Begin VB.TextBox txtfields 
         Alignment       =   1  'Right Justify
         DataField       =   "AR CUST Highest Balance"
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
         Index           =   30
         Left            =   8280
         Locked          =   -1  'True
         TabIndex        =   101
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox txtfields 
         Alignment       =   1  'Right Justify
         DataField       =   "AR CUST Financial Total"
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
         Index           =   29
         Left            =   8160
         Locked          =   -1  'True
         TabIndex        =   100
         Top             =   1800
         Width           =   1815
      End
      Begin VB.TextBox txtfields 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         DataField       =   "AR CUST Financial Period 4"
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
         Index           =   28
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   99
         Top             =   1800
         Width           =   1815
      End
      Begin VB.TextBox txtfields 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000FFFF&
         DataField       =   "AR CUST Financial Period 3"
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
         Index           =   27
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   98
         Top             =   1800
         Width           =   1815
      End
      Begin VB.TextBox txtfields 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF00&
         DataField       =   "AR CUST Financial Period 2"
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
         Index           =   26
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   97
         Top             =   1800
         Width           =   1815
      End
      Begin VB.TextBox txtfields 
         Alignment       =   1  'Right Justify
         DataField       =   "AR CUST Financial Period 1"
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
         Index           =   25
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   96
         Top             =   1800
         Width           =   1815
      End
      Begin VB.TextBox txtfields 
         DataField       =   "AR CUST Name"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   24
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   95
         Top             =   960
         Width           =   3375
      End
      Begin VB.TextBox txtfields 
         DataField       =   "AR CUST Customer ID"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   22
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   94
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Back"
         Height          =   855
         Left            =   9240
         Picture         =   "frm_AR_Customer.frx":49EB
         Style           =   1  'Graphical
         TabIndex        =   88
         Top             =   2160
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Calcula&te"
         Height          =   855
         Left            =   9240
         Picture         =   "frm_AR_Customer.frx":4CF5
         Style           =   1  'Graphical
         TabIndex        =   87
         Top             =   3240
         Width           =   975
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Height          =   1575
         Left            =   2400
         TabIndex        =   140
         Top             =   4200
         Visible         =   0   'False
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   2778
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
            DataField       =   "AGE Sales Doc Ext No"
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
         Alignment       =   2  'Center
         Caption         =   "Last Year"
         DataSource      =   "adoPrimaryRS"
         Height          =   255
         Index           =   42
         Left            =   4320
         TabIndex        =   130
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Caption         =   "Lifetime"
         DataSource      =   "adoPrimaryRS"
         Height          =   255
         Index           =   41
         Left            =   6240
         TabIndex        =   129
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Caption         =   "YTD"
         DataSource      =   "adoPrimaryRS"
         Height          =   255
         Index           =   40
         Left            =   2400
         TabIndex        =   128
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Average Days To Pay:"
         DataSource      =   "adoPrimaryRS"
         Height          =   255
         Index           =   39
         Left            =   6480
         TabIndex        =   127
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Write Offs"
         DataSource      =   "adoPrimaryRS"
         Height          =   255
         Index           =   38
         Left            =   480
         TabIndex        =   126
         Top             =   3720
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Sales"
         DataSource      =   "adoPrimaryRS"
         Height          =   255
         Index           =   37
         Left            =   840
         TabIndex        =   125
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Payments"
         DataSource      =   "adoPrimaryRS"
         Height          =   255
         Index           =   36
         Left            =   720
         TabIndex        =   124
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Invoices"
         DataSource      =   "adoPrimaryRS"
         Height          =   255
         Index           =   34
         Left            =   840
         TabIndex        =   123
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Highest Balance:"
         DataSource      =   "adoPrimaryRS"
         Height          =   255
         Index           =   33
         Left            =   6480
         TabIndex        =   122
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Caption         =   "Total"
         DataSource      =   "adoPrimaryRS"
         Height          =   255
         Index           =   32
         Left            =   8160
         TabIndex        =   121
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Caption         =   "Over 90"
         DataSource      =   "adoPrimaryRS"
         Height          =   255
         Index           =   31
         Left            =   6240
         TabIndex        =   120
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Caption         =   "60 - 90"
         DataSource      =   "adoPrimaryRS"
         Height          =   255
         Index           =   30
         Left            =   4320
         TabIndex        =   119
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Caption         =   "30 - 60"
         DataSource      =   "adoPrimaryRS"
         Height          =   255
         Index           =   29
         Left            =   2400
         TabIndex        =   118
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Caption         =   "0 - 30"
         DataSource      =   "adoPrimaryRS"
         Height          =   255
         Index           =   28
         Left            =   480
         TabIndex        =   117
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Customer Name"
         DataSource      =   "adoPrimaryRS"
         Height          =   255
         Index           =   25
         Left            =   120
         TabIndex        =   116
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Customer ID"
         DataSource      =   "adoPrimaryRS"
         Height          =   255
         Index           =   24
         Left            =   360
         TabIndex        =   115
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Customer Data"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   15
      TabIndex        =   62
      Top             =   120
      Width           =   10335
   End
End
Attribute VB_Name = "frm_AR_Customer"
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

Dim TempStr As String
Dim CurrCust As String
Dim WhichField As String
Dim WhichFields As String

Private Sub btcbRefresh_Click(Index As Integer)
    Dim txttmp As String
    txttmp = cbfields(Index).Text   'saves current value
    loadCombo Index                 'refreshes list (WARNING: destroys current value)
    cbfields(Index).Text = txttmp   'reassigns the same value
End Sub

Private Sub btCustID_Click()
    Dim ghead As String
    Dim fhead As String
        
    ghead = "Customer"
    fhead = "ID//Name"

    AllLookup.ToWhichRecord ADOprimaryrs, ghead, fhead
    'AllLookup.Show vbModal
    

End Sub

Private Sub btCustSalesAcc_Click()
    Dim SQLstatement As String
    Dim No As Integer
    Dim ghead As String
    Dim fhead As String

    No = 1
    SQLstatement = "select [GL COA Account No], [GL COA Account Name]" & _
                    "from [GL Chart Of Accounts]"
    ghead = "Account Description"
    fhead = "Account No//Description"
    
    AllLookup.GetWhichTable No, SQLstatement, ghead, fhead, db
    'AllLookup.Show vbModal
    txtFields(18).SetFocus  ' trigger event adFirstChange
End Sub


Private Sub cbfields_LostFocus(Index As Integer)
Select Case Index
Case 1
   CheckCombo cbfields(Index), "[LIST PAY Description]", "[LIST Payment Terms]", db, True
Case 2
   CheckCombo cbfields(Index), "[Customer Type]", "[LIST Customer Types]", db, True
Case 3
   CheckCombo cbfields(Index), "[SYS TAXGRP ID]", "[SYS Tax Group]", db, True
Case 4
   CheckCombo cbfields(Index), "[EMP ID]", "[EMP Employees]", db, True
End Select
End Sub

Private Sub cdmAgeReceivable_Click()
    AgeSelectedReceivables txtFields(22), db
End Sub

Private Sub cdmShowAll_Click()
    'picButtons.Visible = False
    'frShowAll.Left = 0
    'frShowAll.Top = 480
    Set grdDataGrid.DataSource = ADOprimaryrs
    cmdAdd.Enabled = False
    cmdUpdate.Enabled = False
    cmdDelete.Enabled = False
    cmdRefresh.Enabled = False
    Label1.Caption = "Contact List"
    frShowAll.Visible = True
    frShowAll.ZOrder 0
    'Me.Height = 6750
End Sub

Private Sub cmdAgingDetails_Click()
  Set DataGrid2.DataSource = Nothing
    
  If ADOaging Is Nothing Then
  Else
    ADOaging.Close
    Set ADOaging = Nothing
  End If
  AgingDetails txtFields(22).Text, db
  
  ShowStatus True
    Set ADOaging = New ADODB.Recordset
    ADOaging.Open "SELECT[AGE Sales Doc Ext No],[AGE Start Date],[AGE Orig Amount]," & _
    "[AGE PEriod 1],[AGE Period 2],[AGE PEriod 3],[AGE Period 4] from " & _
    "[AGE AGing Sales Work] WHERE [AGE Cust ID]='" & txtFields(22).Text & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
      
    Dim P1 As Currency, P2 As Currency, P3 As Currency, P4 As Currency, Ptotal As Currency
      
    Do While Not ADOaging.EOF
      P1 = P1 + ADOaging![AGE PEriod 1]
      P2 = P2 + ADOaging![AGE PEriod 2]
      P3 = P3 + ADOaging![AGE PEriod 3]
      P4 = P4 + ADOaging![AGE PEriod 4]
      ADOaging.MoveNext
    Loop
    Ptotal = P1 + P2 + P3 + P4
    ADOprimaryrs![AR CUST Financial Period 1] = FormatCurr(P1)
    ADOprimaryrs![AR CUST Financial Period 2] = FormatCurr(P2)
    ADOprimaryrs![AR CUST Financial Period 3] = FormatCurr(P3)
    ADOprimaryrs![AR CUST Financial Period 4] = FormatCurr(P4)
    ADOprimaryrs![AR CUST Financial Total] = FormatCurr(Ptotal)
    ADOprimaryrs.Update
    
  Set DataGrid2.DataSource = ADOaging
  DataGrid2.Visible = True
  
  ShowStatus False
End Sub

Private Sub cmdAveDays_Click()
    AverageDaysToPay txtFields(22), db
End Sub

Private Sub cmdBack_Click()
If cmdBack.Caption = "Back" Then
    Set grdDataGrid.DataSource = Nothing
    cmdAdd.Enabled = True
    cmdUpdate.Enabled = True
    cmdDelete.Enabled = True
    cmdRefresh.Enabled = True
    Label1.Caption = "Customer Data"
    frShowAll.Visible = False
    Form_Resize
    'Me.Height = 7485
Else
    Unload Me
End If
End Sub

Public Sub ShowList()
    frm_AR_Customer.Show
    cmdBack.Caption = "Close"
    cdmShowAll_Click
End Sub

Private Sub cmdBack1_Click()
    AR_Cust_Drill.Visible = False
    If ADOTransRS Is Nothing Then
    Else
        ADOTransRS.Close
        Set ADOTransRS = Nothing
    End If
    Form_Resize
    cmdAdd.Enabled = True
    cmdUpdate.Enabled = True
    cmdDelete.Enabled = True
    cmdRefresh.Enabled = True
    Label1.Caption = "Customer Data"
End Sub

Private Sub cmdDate_Click(Index As Integer)
Select Case Index
Case 0
    Menu_Calendar.WhoCallMe True, 1645  '17
Case 1
    Menu_Calendar.WhoCallMe True, 1650  '16
End Select
End Sub

Private Sub cmdSearch_Click()
If ADOTransRS Is Nothing Then
Else
    If ADOTransRS.RecordCount = 0 Then Exit Sub
    SearchRECORD ADOTransRS, DataGrid1, txtFields(15).Text, lblLabels(9).Caption, WhichField, "AR SALE Document Type"
End If
End Sub

Private Sub Command4_Click()
If ADOprimaryrs Is Nothing Then
Else
    If ADOprimaryrs.RecordCount = 0 Then Exit Sub
    SearchRECORD ADOprimaryrs, grdDataGrid, txtFields(45).Text, lblLabels(12).Caption, WhichFields, "AR CUST Customer ID"
End If
End Sub

Private Sub cmdShow_Click()
If txtFields(17).Text = "" Or txtFields(16).Text = "" Then
    MsgBox "Please select start and the end date before executing the process"
    Exit Sub
End If
If ADOTransRS Is Nothing Then
Else
    ADOTransRS.Close
    Set ADOTransRS = Nothing
End If
Set ADOTransRS = New ADODB.Recordset
    TempStr = "SELECT[AR SALE Posted YN], [AR SALE Date],[AR SALE Document Type]," & _
    "[AR SALE PO ID],[AR SALE Ext Document #],[AR SALE Payment Terms],[AR SALE Amount Paid]," & _
    "[AR SALE Total] from [AR Sales] WHERE [AR SALE Customer ID]='" & txtFields(0) & "' AND [AR SALE Date] BETWEEN #" & txtFields(17).Text & "# AND #" & txtFields(16).Text & "#"
    
    ADOTransRS.Open TempStr, db, adOpenKeyset, adLockReadOnly, adCmdText
    
    Set DataGrid1.DataSource = ADOTransRS
    If ADOTransRS.RecordCount = 0 Then MsgBox "There is no transaction yet with " & txtFields(0)
End Sub

Private Sub Command1_Click()
If ADOTransRS Is Nothing Then
Else
    ADOTransRS.Close
    Set ADOTransRS = Nothing
End If
Set ADOTransRS = New ADODB.Recordset
    TempStr = "SELECT[AR SALE Posted YN], [AR SALE Date],[AR SALE Document Type]," & _
    "[AR SALE PO ID],[AR SALE Ext Document #],[AR SALE Payment Terms],[AR SALE Amount Paid]," & _
    "[AR SALE Total] from [AR Sales] WHERE [AR SALE Customer ID]='" & txtFields(0) & "'"
    
    ADOTransRS.Open TempStr, db, adOpenKeyset, adLockReadOnly, adCmdText
    
    Set DataGrid1.DataSource = ADOTransRS
    If ADOTransRS.RecordCount = 0 Then
        MsgBox "There is no transaction yet with " & txtFields(0)
    Else
        txtFields(17).Text = ""
        txtFields(16).Text = ""
    End If

End Sub

Private Sub Command2_Click()
  'Recalculate the values on the form
  'Recalc the Total AR Balance
  With ADOprimaryrs
  'MsgBox ![AR CUST Financial Period 1]
  'MsgBox ![AR CUST Financial Period 2]
  'MsgBox ![AR CUST Financial Period 3]
  'MsgBox ![AR CUST Financial Period 4]
  
  'MsgBox CCur(Me.txtFields(28).Text)
  'MsgBox (Me.txtFields(25).Text)
  'MsgBox (Me.txtFields(25).Text)
  'MsgBox CDbl(Me.txtFields(25).Text)
  'MsgBox CCur(Me.txtFields(28).Text)
  'MsgBox CCur(Me.txtFields(27).Text)
  'MsgBox CCur(Me.txtFields(26).Text)
  'MsgBox CCur(Me.txtFields(25).Text)
  ![AR CUST Financial Total] = ![AR CUST Financial Period 1] + ![AR CUST Financial Period 2] + ![AR CUST Financial Period 3] + ![AR CUST Financial Period 4]
  End With
  'Recalc the YTD Fields
  'What is first day of year
  Dim DayOne As Variant
  DayOne = "1/01/" & Format(Now, "yyyy")
  
  Dim WriteOffSource As String
  WriteOffSource = " [AR Payment Header] LEFT JOIN [AR Payment Invoice Cross Reference] ON [AR Payment Header].[AR PAY ID] = [AR Payment Invoice Cross Reference].[AR CROSS Payment ID]"
  
  Sales = 0
  Returns = 0
  gCustomerID$ = Me.txtFields(22).Text
  
  Sales = SumRecord("[AR SALE Total]", "[AR Sales]", db, "[AR SALE Customer ID] = '" & gCustomerID$ & "' AND [AR SALE Posted YN] = TRUE AND [AR SALE Date] >= #" & DayOne & "# AND [AR SALE Document Type] in ('Invoice','Sales Memo','Beginning Balance','Finance Charge')")
  If IsNull(Sales) Then Sales = 0
      
  Returns = SumRecord("[AR SALE Total]", "[AR Sales]", db, "[AR SALE Customer ID] = '" & gCustomerID$ & "' AND [AR SALE Posted YN] = TRUE AND [AR SALE Date] >= #" & DayOne & "# AND [AR SALE Document Type] in ('Return','Credit Memo')")
  If IsNull(Returns) Then Returns = 0
  
  ADOprimaryrs.Fields("AR CUST Sales YTD") = Sales - Returns
  
  Payments = SumRecord("[AR PAY Amount]", "[AR Payment Header]", db, "[AR PAY Customer No] = '" & gCustomerID$ & "' AND [AR PAY Posted YN] = TRUE AND [AR PAY Transaction Date] >= #" & DayOne & "# AND [AR PAY NSF] = FALSE AND [AR PAY Type] <> 'Credit Memo' AND [AR PAY Type] <> 'Return'")
  If IsNull(Payments) Then
      ADOprimaryrs.Fields("AR CUST Payments YTD") = 0
  Else
      ADOprimaryrs.Fields("AR CUST Payments YTD") = Payments
  End If
 
  Invoices = CountRecord("[AR SALE Total]", "[AR Sales]", db, "[AR SALE Customer ID] = '" & gCustomerID$ & "' AND [AR SALE Posted YN] = TRUE AND [AR SALE Date] >= #" & DayOne & "# AND [AR SALE Document Type] in ('Invoice','Sales Memo','Beginning Balance','Finance Charge')")
  If IsNull(Invoices) Then
      ADOprimaryrs.Fields("AR CUST Invoices YTD") = 0
  Else
      ADOprimaryrs.Fields("AR CUST Invoices YTD") = Invoices
  End If
    
  WriteOff = SumRecord("[AR CROSS Write Off Amount]", WriteOffSource, db, "[AR PAY Customer No] = '" & gCustomerID$ & "' AND [AR PAY Transaction Date] >= #" & DayOne & "# AND [AR PAY Posted YN] = TRUE  AND [AR PAY NSF] = FALSE")
  If IsNull(WriteOff) Then
      ADOprimaryrs.Fields("AR CUST Write Offs YTD") = 0
  Else
      ADOprimaryrs.Fields("AR CUST Write Offs YTD") = WriteOff
  End If

  Sales = 0
  Returns = 0
  
  'Last Year
  Dim LastDay As Variant
  LastDay = DateAdd("y", -1, DayOne)
  DayOne = DateAdd("yyyy", -1, DayOne)
  
  Sales = SumRecord("[AR SALE Total]", "[AR Sales]", db, "[AR SALE Customer ID] = '" & gCustomerID$ & "' AND [AR SALE Posted YN] = TRUE AND [AR SALE Date] BETWEEN #" & DayOne & "# AND #" & LastDay & "#  AND [AR SALE Document Type] in ('Invoice','Sales Memo','Beginning Balance','Finance Charge')")
  If IsNull(Sales) Then Sales = 0
  
  Returns = SumRecord("[AR SALE Total]", "[AR Sales]", db, "[AR SALE Customer ID] = '" & gCustomerID$ & "' AND [AR SALE Posted YN] = TRUE AND [AR SALE Date] BETWEEN #" & DayOne & "# AND #" & LastDay & "#  AND [AR SALE Document Type] in ('Return','Credit Memo')")
    If IsNull(Returns) Then Returns = 0
  
  ADOprimaryrs.Fields("AR CUST Sales Last Year") = Sales - Returns
  
  Payments = SumRecord("[AR PAY Amount]", "[AR Payment Header]", db, "[AR PAY Customer No] = '" & gCustomerID$ & "' AND [AR PAY Posted YN] = TRUE AND [AR PAY Transaction Date] BETWEEN #" & DayOne & "# AND #" & LastDay & "# AND [AR PAY NSF] = FALSE AND [AR PAY Type] <> 'Credit Memo' AND [AR PAY Type] <> 'Return'")
  If IsNull(Payments) Then
    ADOprimaryrs.Fields("AR CUST Payments Last Year") = 0
  Else
    ADOprimaryrs.Fields("AR CUST Payments Last Year") = Payments
  End If
 
  Invoices = CountRecord("[AR SALE Total]", "[AR Sales]", db, "[AR SALE Customer ID] = '" & gCustomerID$ & "' AND [AR SALE Posted YN] = TRUE AND [AR SALE Date] BETWEEN #" & DayOne & "# AND #" & LastDay & "# AND [AR SALE Document Type] in ('Invoice','Sales Memo','Beginning Balance','Finance Charge')")
  If IsNull(Invoices) Then
    ADOprimaryrs.Fields("AR CUST Invoices Last Year") = 0
  Else
    ADOprimaryrs.Fields("AR CUST Invoices Last Year") = Invoices
 End If

  WriteOff = SumRecord("[AR CROSS Write Off Amount]", WriteOffSource, db, "[AR PAY Customer No] = '" & gCustomerID$ & "' AND [AR PAY Transaction Date] BETWEEN #" & DayOne & "# AND #" & LastDay & "# AND [AR PAY Posted YN] = TRUE  AND [AR PAY NSF] = FALSE")
  If IsNull(WriteOff) Then
    ADOprimaryrs.Fields("AR CUST Write Offs Last Year") = 0
  Else
    ADOprimaryrs.Fields("AR CUST Write Offs Last Year") = WriteOff
  End If

  Sales = 0
  Returns = 0

  'Lifetime
  Sales = SumRecord("[AR SALE Total]", "[AR Sales]", db, "[AR SALE Customer ID] = '" & gCustomerID$ & "' AND [AR SALE Posted YN] = TRUE AND [AR SALE Document Type] in ('Invoice','Sales Memo','Beginning Balance','Finance Charge')")
  If IsNull(Sales) Then Sales = 0
  
  Returns = SumRecord("[AR SALE Total]", "[AR Sales]", db, "[AR SALE Customer ID] = '" & gCustomerID$ & "' AND [AR SALE Posted YN] = TRUE AND [AR SALE Document Type] in ('Return','Credit Memo')")
  If IsNull(Returns) Then Returns = 0
  
  ADOprimaryrs.Fields("AR CUST Sales Lifetime") = Sales - Returns
  
  Payments = SumRecord("[AR PAY Amount]", "[AR Payment Header]", db, "[AR PAY Customer No] = '" & gCustomerID$ & "' AND [AR PAY Posted YN] = TRUE AND [AR PAY NSF] = FALSE AND [AR PAY Type] <> 'Credit Memo' AND [AR PAY Type] <> 'Return'")
  If IsNull(Payments) Then
    ADOprimaryrs.Fields("AR CUST Payments Lifetime") = 0
  Else
    ADOprimaryrs.Fields("AR CUST Payments Lifetime") = Payments
  End If
      
  Invoices = CountRecord("[AR SALE Total]", "[AR Sales]", db, "[AR SALE Customer ID] = '" & gCustomerID$ & "' AND [AR SALE Posted YN] = TRUE AND [AR SALE Document Type] in ('Invoice','Sales Memo','Beginning Balance','Finance Charge')")
    If IsNull(Invoices) Then
    ADOprimaryrs.Fields("AR CUST Invoices Lifetime") = 0
  Else
    ADOprimaryrs.Fields("AR CUST Invoices Lifetime") = Invoices
  End If
    
  WriteOff = SumRecord("[AR CROSS Write Off Amount]", WriteOffSource, db, "[AR PAY Customer No] = '" & gCustomerID$ & "' AND [AR PAY Posted YN] = TRUE  AND [AR PAY NSF] = FALSE")
  If IsNull(WriteOff) Then
    ADOprimaryrs.Fields("AR CUST Write Offs Lifetime") = 0
  Else
    ADOprimaryrs.Fields("AR CUST Write Offs Lifetime") = WriteOff
  End If
    
ADOprimaryrs.UpdateBatch adAffectAll
'ADOprimaryrs.Requery
End Sub

Private Sub Command3_Click()

    Set DataGrid2.DataSource = Nothing
    If ADOaging Is Nothing Then
    Else
        ADOaging.Close
        Set ADOaging = Nothing
        DataGrid2.Visible = False
    End If
    
    AR_Cust_Financials.Visible = False
    Form_Resize
    cmdAdd.Enabled = True
    cmdUpdate.Enabled = True
    cmdDelete.Enabled = True
    cmdRefresh.Enabled = True
    Label1.Caption = "Customer Data"
End Sub


Private Sub CmdPrtState_Click()
    BuildStatement txtFields(22), db
End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
If ADOTransRS.RecordCount = 0 Then Exit Sub
    lblLabels(9) = DataGrid1.Columns(ColIndex).Caption
    WhichField = DataGrid1.Columns(ColIndex).DataField
    ADOTransRS.Close
    Set ADOTransRS = Nothing
    Set ADOTransRS = New ADODB.Recordset
    ADOTransRS.Open TempStr & " ORDER BY [" & WhichField & "]", db, adOpenKeyset, adLockOptimistic, adCmdText
    Set DataGrid1.DataSource = ADOTransRS
End Sub

Private Sub FinBut_Click()
    'gCustomerdrill = frm_AR_Customer.txtFields(0).Text
    'frm_AR_Cust_Financials.Show
    
    DataGrid2.Columns(3).Caption = "0 to " & LookRecord("[SYS COM Sales Period 1]", "[SYS Company]", db) & " Days"
    DataGrid2.Columns(4).Caption = LookRecord("[SYS COM Sales Period 1]", "[SYS Company]", db) + 1 & " To " & LookRecord("[SYS COM Sales Period 2]", "[SYS Company]", db) & " Days"
    DataGrid2.Columns(5).Caption = LookRecord("[SYS COM Sales Period 2]", "[SYS Company]", db) + 1 & " To " & LookRecord("[SYS COM Sales Period 3]", "[SYS Company]", db) & " Days"
    DataGrid2.Columns(6).Caption = "Over " & LookRecord("[SYS COM Sales Period 3]", "[SYS Company]", db) & " Days"
    
    AR_Cust_Financials.ZOrder 0
    AR_Cust_Financials.Visible = True
    Form_Resize
    cmdAdd.Enabled = False
    cmdUpdate.Enabled = False
    cmdDelete.Enabled = False
    cmdRefresh.Enabled = False
    Label1.Caption = "Customer Finacial Report"
End Sub

Private Sub Form_Load()
mbAddNewFlag = False
ShowStatus True
On Error GoTo FormErr
  
  Me.Width = 10485
  Me.Height = 7500

    Set db = New ADODB.Connection
    db.CursorLocation = adUseClient
    db.Open gblADOProvider

  Dim sql As String
  sql = "select [AR CUST Customer ID],[AR CUST Name],[AR CUST Address 1]," & _
    "[AR CUST Address 2],[AR CUST City],[AR CUST State],[AR CUST Postal]," & _
    "[AR CUST Country],[AR CUST Contact],[AR CUST Phone],[AR CUST Phone Ext]," & _
    "[AR CUST Fax],[AR CUST Tax ID No],[AR CUST Average Days To Pay]," & _
    "[AR CUST Credit Limit],[AR CUST Payment Terms],[AR CUST Type]," & _
    "[AR CUST Tax Group],[AR CUST Sales Account],[AR CUST SalesPerson]," & _
    "[AR CUST Discount %],[AR CUST Statements YN],[AR CUST Finance Charge YN]," & _
    "[AR CUST Sales YTD],[AR CUST Notes],[AR CUST Web Page],[AR EMail Address]," & _
    "[AR CUST Financial Period 2],[AR CUST Financial Period 3],[AR CUST Financial Period 4]," & _
    "[AR CUST Financial Period 1],[AR CUST Financial Total],[AR CUST Sales Last Year]," & _
    "[AR CUST Sales Lifetime],[AR CUST Payments YTD],[AR CUST Payments Last Year]," & _
    "[AR CUST Payments Lifetime],[AR CUST Write Offs YTD],[AR CUST Write Offs Last Year]," & _
    "[AR CUST Write Offs Lifetime],[AR CUST Invoices YTD],[AR CUST Invoices Last Year]," & _
    "[AR CUST Invoices Lifetime],[AR CUST Highest Balance],[AR CUST Average Days To Pay] " & _
    "from [AR Customer]"
  Set ADOprimaryrs = New ADODB.Recordset
  ADOprimaryrs.Open sql, db, adOpenStatic, adLockOptimistic
  
  If ADOprimaryrs.RecordCount = 0 Then
    cmdUpdate.Enabled = False
    cmdDelete.Enabled = False
    cmdRefresh.Enabled = False
    ShipToBut.Enabled = False
    FinBut.Enabled = False
    Tranbut.Enabled = False
    btCustSalesAcc.Enabled = False
    btCustID.Enabled = False
  End If
  loadCombo 0   ' loads all comboboxes
  
  Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtFields
    If oText.DataField <> "" Then
        Set oText.DataSource = ADOprimaryrs
        If ADOprimaryrs("" & oText.DataField & "").Type = 202 Then oText.MaxLength = ADOprimaryrs("" & oText.DataField & "").DefinedSize
    End If
  Next

  Dim oCheck As CheckBox
    'Bind the Check boxes to the data provider
  For Each oCheck In Me.chkFields
    Set oCheck.DataSource = ADOprimaryrs
  Next

  Dim oCombo As ComboBox
  For Each oCombo In Me.cbfields
    Set oCombo.DataSource = ADOprimaryrs
  Next
  
  If CheckNewDB(ADOprimaryrs, "Company") = True Then
    cmdAdd_Click
  End If
  
  'Set grdDataGrid.DataSource = ADOprimaryrs
  
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
  
  Me.Width = 10485
  Me.Height = 7500
  
SkipResize:
  frPrimary.Left = (Me.ScaleWidth - frPrimary.Width) / 2
  Label1.Left = frPrimary.Left
  Label1.Width = frPrimary.Width
  frPrimary.Top = (Me.ScaleHeight - frPrimary.Height - picButtons.Height) / 2 + 230
  AR_Cust_Drill.Left = frPrimary.Left
  AR_Cust_Drill.Top = frPrimary.Top
  AR_Cust_Financials.Left = frPrimary.Left
  AR_Cust_Financials.Top = frPrimary.Top
  frShowAll.Top = frPrimary.Top
  frShowAll.Left = frPrimary.Left
  
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
    Case vbKeyPageUp
      If Shift = vbCtrlMask Then
        cmdFirst_Click
      Else
        cmdPrevious_Click
      End If
    Case vbKeyPageDown
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
  ShowStatus True
  
      EndLoad db, ADOprimaryrs, "Customers"
      If ADOprimaryrs.RecordCount > 0 Then
        'If ADOprimaryrs.EditMode <> 0 Then
          ADOprimaryrs.CancelUpdate
        'End If
      End If
      ADOprimaryrs.Close
      Set ADOprimaryrs = Nothing
FormErr:
      db.Close
      Set db = Nothing
  ShowStatus False
  Set frm_AR_Customer = Nothing
End Sub

Private Sub adoPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This will display the current record position for this recordset
  lblStatus.Caption = "Record: " & CStr(ADOprimaryrs.AbsolutePosition) & " of " & CStr(ADOprimaryrs.RecordCount)
  If ADOprimaryrs.BOF Or ADOprimaryrs.EOF Then Exit Sub
  CurrCust = ADOprimaryrs![AR CUST Customer ID] & ""
  txtFields(34).Text = CurrCust
  DataGrid2.Visible = False
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
Dim SetFlagBool As Boolean
  'On Error GoTo AddErr
  With ADOprimaryrs
    If cmdAdd.Caption = "&Add" Then
        'Set grdDataGrid.DataSource = Nothing
        If Not (.BOF And .EOF) And Not mbAddNewFlag Then
          mvBookMark = .Bookmark
        End If
        .AddNew
        lblStatus.Caption = "Add record"
        mbAddNewFlag = True
        'txtFields(0).Enabled = True
'        txtfields(0).SetFocus

            Dim cbCheck As ComboBox
              'Bind the Check boxes to the data provider
            For Each cbCheck In Me.cbfields
              cbCheck.Text = cbCheck.List(0)
            Next

        SetFlagBool = False
        cmdAdd.Caption = "&Cancel"
        btCustSalesAcc.Enabled = True
        cmdUpdate.Enabled = True
        'txtFields(0).Locked = False
    Else
        mbAddNewFlag = False
        cmdCancel_Click
        cmdAdd.Caption = "&Add"
        SetFlagBool = True
        'txtFields(0).Enabled = False
        'txtFields(0).Locked = True
        'Set grdDataGrid.DataSource = ADOprimaryrs
    End If
  End With
  
        btCustID.Enabled = SetFlagBool
        cmdDelete.Enabled = SetFlagBool
        cmdRefresh.Enabled = SetFlagBool
        ShipToBut.Enabled = SetFlagBool
        FinBut.Enabled = SetFlagBool
        Tranbut.Enabled = SetFlagBool
  
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
            ShipToBut.Enabled = False
            FinBut.Enabled = False
            Tranbut.Enabled = False
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
  'On Error GoTo RefreshErr
    With ADOprimaryrs
        If .EditMode = adEditInProgress Then .CancelUpdate
        If .RecordCount > 0 Then
            mvBookMark = .Bookmark
            .Requery
            .Bookmark = mvBookMark
        End If
    End With
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub
Private Sub cmdCancel_Click()
  'On Error Resume Next

  mbEditFlag = False
  mbAddNewFlag = False
  With ADOprimaryrs
    .CancelUpdate
    If (.BOF And .EOF) Then Exit Sub  ' if no records, don't move
    If mvBookMark > 0 Then
        .Bookmark = mvBookMark
    Else
        .MoveFirst
    End If
  End With
  mbDataChanged = False

End Sub

Private Sub cmdUpdate_Click()
'    On Error GoTo UpdateErr

    With ADOprimaryrs
        If .RecordCount = 0 Then Exit Sub 'no records to update
        If Trim(txtFields(0).Text) <> "" Then
        Dim oTxt As TextBox
          For Each oTxt In Me.txtFields
            If oTxt.Text = "" And oTxt.DataField <> "" Then
              If ADOprimaryrs("" & oTxt.DataField & "").Type = 203 Or ADOprimaryrs("" & oTxt.DataField & "").Type = 202 Then oTxt.Text = " "
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
        txtFields(0).Enabled = True
        btCustID.Enabled = True
        cmdDelete.Enabled = True
        cmdRefresh.Enabled = True
        ShipToBut.Enabled = True
        FinBut.Enabled = True
        Tranbut.Enabled = True
        txtFields(0).Locked = True
    End With

  mbEditFlag = False
  GetTextColor Me
  'Set grdDataGrid.DataSource = ADOprimaryrs
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

Private Sub grdDataGrid_HeadClick(ByVal ColIndex As Integer)
If ADOprimaryrs.RecordCount = 0 Then Exit Sub
    lblLabels(12) = grdDataGrid.Columns(ColIndex).Caption
    WhichFields = grdDataGrid.Columns(ColIndex).DataField
    ADOprimaryrs.Close
    Set ADOprimaryrs = Nothing
'    Set ADOprimaryrs = New ADODB.Recordset
'    ADOTransRS.Open TempStr & " ORDER BY [" & WhichField & "]", db, adOpenKeyset, adLockOptimistic, adCmdText
'    Set DataGrid1.DataSource = ADOTransRS
    
'        If grdDataGrid.Columns(ColIndex).DataField <> "AR CUST Notes" Then
'            ADOprimaryrs.Close
            Set ADOprimaryrs = New ADODB.Recordset
            ADOprimaryrs.Open "select [AR CUST Customer ID],[AR CUST Name],[AR CUST Contact],[AR CUST Phone],[AR CUST Phone Ext],[AR CUST Sales YTD],[AR CUST Notes] from [AR Customer] Order by [" & WhichFields & "]", db, adOpenStatic, adLockOptimistic, adCmdText
            Set grdDataGrid.DataSource = ADOprimaryrs
'        Else
'            ADOprimaryrs.Close
'            Set ADOprimaryrs = New ADODB.Recordset
'            ADOprimaryrs.Open "select [AR CUST Customer ID],[AR CUST Name],[AR CUST Contact],[AR CUST Phone],[AR CUST Phone Ext],[AR CUST Sales YTD],[AR CUST Notes] from [AR Customer] Order by [" & grdDataGrid.Columns((ColIndex) - 1).DataField & "]", db, adOpenStatic, adLockOptimistic
'            Set grdDataGrid.DataSource = ADOprimaryrs
'        End If
'        ADOprimaryrs.Requery
End Sub

Private Sub ShipToBut_Click()
    frm_AR_Cust_Ship_To.CallByUserShip txtFields(0), True
End Sub

Private Sub Tranbut_Click()
    AR_Cust_Drill.Visible = True
    AR_Cust_Drill.ZOrder 0
    cmdAdd.Enabled = False
    cmdUpdate.Enabled = False
    cmdDelete.Enabled = False
    cmdRefresh.Enabled = False
    txtFields(17).Text = ""
    txtFields(16).Text = ""
    Label1.Caption = "Customer Transaction"
    Form_Resize
End Sub

Private Sub loadCombo(cbType As Integer)
    Select Case cbType
    Case 0
        ComboInit cbfields(1), lblfields(1), "select [LIST PAY Description] " & _
            "from [LIST Payment Terms]"
        ComboInit cbfields(2), lblfields(2), "select [Customer Type] " & _
            "from [LIST Customer Types]"
        ComboInit cbfields(3), lblfields(3), "select [SYS TAXGRP ID] from [SYS Tax Group]"
        ComboInit cbfields(4), lblfields(4), "select [EMP ID] from [EMP Employees]"
    Case 1
        ComboInit cbfields(1), lblfields(1), "select [LIST PAY Description] " & _
            "from [LIST Payment Terms]"
    Case 2
        ComboInit cbfields(2), lblfields(2), "select [Customer Type] " & _
            "from [LIST Customer Types]"
    Case 3
        ComboInit cbfields(3), lblfields(3), "select [SYS TAXGRP ID] from [SYS Tax Group]"
    Case 4
        ComboInit cbfields(4), lblfields(4), "select [EMP ID] from [EMP Employees]"
    End Select
End Sub

Private Sub txtfields_GotFocus(Index As Integer)
    TxtGotFocus txtFields(Index)
End Sub

Private Sub txtfields_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
Case 13, 14, 20
    keyResponse = CtrlValidate(KeyAscii, "0123456789.")
    If keyResponse = True Then
    Else
       KeyAscii = 0
    End If
Case 18
    keyResponse = CtrlValidate(KeyAscii, "0123456789")
    If keyResponse = True Then
    Else
       KeyAscii = 0
    End If
End Select
End Sub

Private Sub txtFields_LostFocus(Index As Integer)
Select Case Index
Case 34
    Dim txtCustID As String
    txtCustID = txtFields(34)
    If txtFields(34) = "" And mbAddNewFlag = False Then
        txtFields(34) = CurrCust
    ElseIf txtFields(34) <> "" And mbAddNewFlag = True Then
        If CheckDocument("SELECT [AR CUST Customer ID] FROM [AR Customer] WHERE [AR CUST Customer ID]='" & txtCustID & "'", db, False) = False Then
            MsgBox txtCustID & " is already exist", vbInformation, "Information"
            txtFields(34) = ""
            txtFields(0) = txtFields(34)
            Exit Sub
        Else
            txtFields(0) = txtCustID
        End If
    End If
    
    If txtFields(34) = CurrCust Then Exit Sub
    
    With ADOprimaryrs
      If .RecordCount > 0 And mbAddNewFlag = False Then
        mvBookMark = .Bookmark
        .MoveFirst
        .Find "[AR CUST Customer ID]='" & txtCustID & "'"
        If .EOF Then
            Dim Response As Integer
            Response = MsgBox(txtCustID & " is a new input. Would you like to add it into the database", vbYesNo, "Information")
            If Response = vbYes Then
                mbAddNewFlag = True
                cmdAdd_Click
                txtFields(34) = txtCustID
            Else
                .Bookmark = mvBookMark
                txtFields(34) = txtFields(0)
            End If
            txtFields(34).SetFocus
        End If
      End If
    'Else
    End With
Case 13
    txtFields(Index) = Val(txtFields(Index))
Case 14
    If txtFields(Index) = "" Then txtFields(Index) = 0
    txtFields(Index) = FormatCurr(txtFields(Index))
Case 18
    If Trim(txtFields(Index).Text) = "" Then Exit Sub
    If IsNumeric(txtFields(18).Text) Then
        CheckDocument "SELECT [GL COA Account No] FROM [GL Chart Of Accounts] WHERE [GL COA Account No]='" & txtFields(18).Text & "'", db, False, txtFields(18)
    Else
        MsgBox "Only numeric character is accepted", vbInformation, "Information"
        txtFields(18) = ""
    End If
Case 20
    txtFields(Index) = Format(txtFields(Index), "00.00")
End Select
End Sub

Public Sub CallByUserCust(CustID As String)
    Me.Show
    If mbAddNewFlag = False Then
        cmdAdd_Click
    End If
    txtFields(34) = CustID
    txtFields(0) = CustID
End Sub

