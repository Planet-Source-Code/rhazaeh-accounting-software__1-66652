VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_AR_Cust_Projects 
   Caption         =   "Projects"
   ClientHeight    =   7050
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   9780
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7050
   ScaleWidth      =   9780
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   9780
      TabIndex        =   37
      Top             =   6450
      Width           =   9780
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   300
         Left            =   4320
         TabIndex        =   31
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   300
         Left            =   3240
         TabIndex        =   30
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   300
         Left            =   2160
         TabIndex        =   29
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   300
         Left            =   1080
         TabIndex        =   28
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   300
         Left            =   0
         TabIndex        =   27
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
      ScaleWidth      =   9780
      TabIndex        =   0
      Top             =   6750
      Width           =   9780
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Picture         =   "frm_AR_Cust_Projects.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Picture         =   "frm_AR_Cust_Projects.frx":0342
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Picture         =   "frm_AR_Cust_Projects.frx":0684
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Picture         =   "frm_AR_Cust_Projects.frx":09C6
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   36
         Top             =   0
         Width           =   3360
      End
   End
   Begin VB.Frame frPrimary 
      Height          =   5895
      Left            =   0
      TabIndex        =   38
      Top             =   480
      Width           =   9735
      Begin VB.CommandButton cmdComplete 
         Caption         =   "Complete"
         Height          =   900
         Left            =   6360
         Picture         =   "frm_AR_Cust_Projects.frx":0D08
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   4920
         Width           =   975
      End
      Begin VB.TextBox txtFields 
         DataField       =   "PROJ ID"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   0
         Left            =   1800
         TabIndex        =   1
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdProjectID 
         Height          =   285
         Left            =   3000
         Picture         =   "frm_AR_Cust_Projects.frx":1012
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   960
         Width           =   375
      End
      Begin VB.CommandButton TransButton 
         Caption         =   "Transaction"
         Height          =   900
         Left            =   7440
         Picture         =   "frm_AR_Cust_Projects.frx":115C
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   4920
         Width           =   975
      End
      Begin VB.CommandButton ReCalcButton 
         Caption         =   "Calculate"
         Height          =   900
         Left            =   8520
         Picture         =   "frm_AR_Cust_Projects.frx":1466
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   4920
         Width           =   975
      End
      Begin VB.PictureBox picMajor 
         BorderStyle     =   0  'None
         Height          =   5055
         Left            =   480
         ScaleHeight     =   5055
         ScaleWidth      =   8775
         TabIndex        =   55
         Top             =   720
         Width           =   8775
         Begin VB.CommandButton cmdDate 
            Height          =   285
            Index           =   12
            Left            =   3120
            Picture         =   "frm_AR_Cust_Projects.frx":1770
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   1800
            Width           =   375
         End
         Begin VB.CommandButton cmdDate 
            Height          =   285
            Index           =   4
            Left            =   3120
            Picture         =   "frm_AR_Cust_Projects.frx":1D4A
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   1440
            Width           =   375
         End
         Begin VB.CommandButton btcbRefresh 
            Height          =   280
            Left            =   7560
            Picture         =   "frm_AR_Cust_Projects.frx":2324
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Update the Ship Via"
            Top             =   1800
            Width           =   375
         End
         Begin VB.CommandButton btCust 
            Height          =   285
            Left            =   7560
            Picture         =   "frm_AR_Cust_Projects.frx":27DA
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   240
            Width           =   375
         End
         Begin VB.ComboBox cbfields 
            DataField       =   "PROJ Type"
            Height          =   315
            Left            =   5760
            TabIndex        =   12
            Top             =   1800
            Width           =   1815
         End
         Begin VB.TextBox txtFields 
            DataField       =   "PROJ Name"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   1
            Left            =   1320
            TabIndex        =   3
            Top             =   720
            Width           =   2775
         End
         Begin VB.TextBox txtFields 
            DataField       =   "PROJ Notes"
            DataSource      =   "adoPrimaryRS"
            Height          =   885
            Index           =   2
            Left            =   4560
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   14
            Top             =   2520
            Width           =   3975
         End
         Begin VB.TextBox txtFields 
            DataField       =   "PROJ Project Manager"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   3
            Left            =   1320
            TabIndex        =   4
            Top             =   1080
            Width           =   2775
         End
         Begin VB.TextBox txtFields 
            DataField       =   "PROJ Start Date"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "MM/dd/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   4
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   1440
            Width           =   1815
         End
         Begin VB.TextBox txtFields 
            DataField       =   "PROJ Customer ID"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   6
            Left            =   5760
            TabIndex        =   10
            Top             =   240
            Width           =   1815
         End
         Begin VB.TextBox txtFields 
            DataField       =   "PROJ Description"
            DataSource      =   "adoPrimaryRS"
            Height          =   885
            Index           =   7
            Left            =   0
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   9
            Top             =   2520
            Width           =   4095
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            DataField       =   "PROJ Estimated Cost"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   2
            EndProperty
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   8
            Left            =   960
            TabIndex        =   15
            Top             =   3960
            Width           =   1335
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            DataField       =   "PROJ Estimated Revenue"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   2
            EndProperty
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   9
            Left            =   960
            TabIndex        =   18
            Top             =   4320
            Width           =   1335
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            DataField       =   "PROJ Actual Cost"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   2
            EndProperty
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   10
            Left            =   2400
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   3960
            Width           =   1335
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            DataField       =   "PROJ Actual Revenue"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   2
            EndProperty
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   11
            Left            =   2400
            Locked          =   -1  'True
            TabIndex        =   19
            Top             =   4320
            Width           =   1335
         End
         Begin VB.TextBox txtFields 
            DataField       =   "PROJ Completion Date"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "MM/dd/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   12
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   1800
            Width           =   1815
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   2
            EndProperty
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   15
            Left            =   3960
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   3960
            Width           =   1335
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   2
            EndProperty
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   21
            Left            =   3960
            Locked          =   -1  'True
            TabIndex        =   20
            Top             =   4320
            Width           =   1335
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   2
            EndProperty
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   20
            Left            =   960
            TabIndex        =   21
            Top             =   4680
            Width           =   1335
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   2
            EndProperty
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   18
            Left            =   2400
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   4680
            Width           =   1335
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   2
            EndProperty
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   19
            Left            =   3960
            Locked          =   -1  'True
            TabIndex        =   23
            Top             =   4680
            Width           =   1335
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Project ID"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   74
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Customer Phone"
            DataSource      =   "adoPrimaryRS"
            Height          =   255
            Index           =   16
            Left            =   4320
            TabIndex        =   73
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Customer Name"
            DataSource      =   "adoPrimaryRS"
            Height          =   255
            Index           =   15
            Left            =   4440
            TabIndex        =   72
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label lblfields 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   1
            Left            =   5760
            TabIndex        =   71
            Top             =   1080
            Width           =   2175
         End
         Begin VB.Label lblfields 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   0
            Left            =   5760
            TabIndex        =   70
            Top             =   720
            Width           =   2175
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "End Date"
            Height          =   255
            Index           =   14
            Left            =   240
            TabIndex        =   69
            Top             =   1800
            Width           =   975
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Project Name"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   68
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Notes"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   4560
            TabIndex        =   67
            Top             =   2280
            Width           =   3975
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Project Manager"
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   66
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Start Date"
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   65
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Project Type"
            Height          =   255
            Index           =   5
            Left            =   4680
            TabIndex        =   64
            Top             =   1800
            Width           =   975
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Customer ID"
            DataSource      =   "adoPrimaryRS"
            Height          =   255
            Index           =   6
            Left            =   4680
            TabIndex        =   63
            Top             =   240
            Width           =   975
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Project Description"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   7
            Left            =   0
            TabIndex        =   62
            Top             =   2280
            Width           =   4095
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Cost"
            Height          =   255
            Index           =   8
            Left            =   0
            TabIndex        =   61
            Top             =   3960
            Width           =   855
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Revenue"
            Height          =   255
            Index           =   9
            Left            =   0
            TabIndex        =   60
            Top             =   4320
            Width           =   855
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            Caption         =   "Estimated"
            DataSource      =   "adoPrimaryRS"
            Height          =   255
            Index           =   12
            Left            =   960
            TabIndex        =   59
            Top             =   3720
            Width           =   1335
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            Caption         =   "Actual"
            DataSource      =   "adoPrimaryRS"
            Height          =   255
            Index           =   13
            Left            =   2400
            TabIndex        =   58
            Top             =   3720
            Width           =   1335
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            Caption         =   "Variance"
            DataSource      =   "adoPrimaryRS"
            Height          =   255
            Index           =   10
            Left            =   3960
            TabIndex        =   57
            Top             =   3720
            Width           =   1335
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Net"
            Height          =   255
            Index           =   11
            Left            =   0
            TabIndex        =   56
            Top             =   4680
            Width           =   855
         End
      End
      Begin VB.Image Image1 
         Height          =   600
         Left            =   8040
         Picture         =   "frm_AR_Cust_Projects.frx":2924
         Top             =   120
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Closed On 25/04/2000"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   600
         TabIndex        =   75
         Top             =   360
         Visible         =   0   'False
         Width           =   1650
      End
   End
   Begin VB.Frame AR_Cust_Proj_Drill 
      Height          =   5895
      Left            =   0
      TabIndex        =   40
      Top             =   480
      Width           =   9735
      Begin VB.Frame Frame3 
         Height          =   1335
         Left            =   4200
         TabIndex        =   49
         Top             =   120
         Width           =   2895
         Begin VB.CommandButton cmdSearch 
            Caption         =   "&Search"
            Height          =   855
            Left            =   1680
            Picture         =   "frm_AR_Cust_Projects.frx":3500
            Style           =   1  'Graphical
            TabIndex        =   51
            Top             =   360
            Width           =   975
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
            TabIndex        =   50
            Top             =   840
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
            Index           =   17
            Left            =   240
            TabIndex        =   52
            Top             =   600
            Width           =   1335
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1335
         Left            =   120
         TabIndex        =   43
         Top             =   120
         Width           =   3975
         Begin VB.CommandButton Command1 
            Caption         =   "All Project"
            Height          =   855
            Left            =   3960
            Picture         =   "frm_AR_Cust_Projects.frx":380A
            Style           =   1  'Graphical
            TabIndex        =   76
            Top             =   360
            Width           =   975
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
            Index           =   16
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   54
            Top             =   840
            Width           =   1455
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
            Index           =   17
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   53
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton cmdShow 
            Caption         =   "&Execute"
            Height          =   855
            Left            =   2880
            Picture         =   "frm_AR_Cust_Projects.frx":3B14
            Style           =   1  'Graphical
            TabIndex        =   46
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton cmdDate 
            Height          =   285
            Index           =   1
            Left            =   2400
            Picture         =   "frm_AR_Cust_Projects.frx":3E1E
            Style           =   1  'Graphical
            TabIndex        =   45
            Top             =   840
            Width           =   375
         End
         Begin VB.CommandButton cmdDate 
            Height          =   285
            Index           =   0
            Left            =   2400
            Picture         =   "frm_AR_Cust_Projects.frx":43F8
            Style           =   1  'Graphical
            TabIndex        =   44
            Top             =   360
            Width           =   375
         End
         Begin VB.Label lblLabels 
            Caption         =   "End Date:"
            Height          =   255
            Index           =   26
            Left            =   120
            TabIndex        =   48
            Top             =   840
            Width           =   975
         End
         Begin VB.Label lblLabels 
            Caption         =   "Start Date:"
            Height          =   255
            Index           =   27
            Left            =   120
            TabIndex        =   47
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "Back"
         Height          =   855
         Left            =   8520
         Picture         =   "frm_AR_Cust_Projects.frx":49D2
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   480
         Width           =   975
      End
      Begin MSDataGridLib.DataGrid grdDataGrid 
         Height          =   4215
         Left            =   120
         TabIndex        =   41
         Top             =   1560
         Width           =   9465
         _ExtentX        =   16695
         _ExtentY        =   7435
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
         Caption         =   "Project Transaction"
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
         BeginProperty Column02 
            DataField       =   "AR SALE Salesperson"
            Caption         =   "Salesperson"
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
            Caption         =   "PO No."
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
         BeginProperty Column05 
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
         BeginProperty Column06 
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
         BeginProperty Column07 
            DataField       =   "AR SALE Total"
            Caption         =   "Total"
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
               ColumnWidth     =   659.906
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1244.976
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1049.953
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1124.787
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1170.142
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Project Information"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   39
      Top             =   120
      Width           =   9135
   End
End
Attribute VB_Name = "frm_AR_Cust_Projects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ADOprimaryrs As ADODB.Recordset
Attribute ADOprimaryrs.VB_VarHelpID = -1
Dim ADOTransRS As ADODB.Recordset

Dim db As ADODB.Connection
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean

Dim TempStr As String
Dim WhichField As String

Private Sub btcbRefresh_Click()
    Dim tmp As String
    tmp = cbfields.Text
    loadCombo
    cbfields.Text = tmp
End Sub

Private Sub btCust_Click()
    Dim sql As String
    Dim No As Integer
    Dim ghead As String
    Dim fhead As String
    
    No = 37
    sql = "select [AR CUST Customer ID],[AR CUST Name] from [AR Customer]"
    ghead = "Customer Information"
    fhead = "ID//Name"
    
    AllLookup.GetWhichTable No, sql, ghead, fhead, db
    'AllLookup.Show vbModal
    txtfields(6).SetFocus 'trigger event adFirstChange
    'updates the additional information textboxes
    If txtfields(6).Text <> "" Then
        lblfields(0) = LookRecord("[AR CUST Name]", "[AR Customer]", db, "[AR CUST Customer ID] = '" & Me.txtfields(6).Text & "'")
        lblfields(1) = LookRecord("[AR CUST Phone]", "[AR Customer]", db, "[AR CUST Customer ID] = '" & Me.txtfields(6).Text & "'") & " Ext. " & LookRecord("[AR CUST Phone Ext]", "[AR Customer]", db, "[AR CUST Customer ID] = '" & Me.txtfields(6).Text & "'")
    End If
End Sub

Private Sub cbfields_KeyPress(KeyAscii As Integer)
    keyResponse = CtrlValidate(KeyAscii, "")
    If keyResponse = True Then
    Else
       KeyAscii = 0
    End If
End Sub

Private Sub cbfields_LostFocus()
   If CheckCombo(cbfields) Then
        MsgBox "There is no such selection", vbInformation, "Information"
   End If
End Sub

Private Sub cmdBack_Click()
    If ADOTransRS Is Nothing Then
    Else
        ADOTransRS.Close
        Set ADOTransRS = Nothing
    End If
    cmdAdd.Enabled = True
    cmdUpdate.Enabled = True
    cmdDelete.Enabled = True
    cmdRefresh.Enabled = True
    frPrimary.Visible = True
    AR_Cust_Proj_Drill.Visible = False
    Form_Resize
End Sub

Private Sub cmdComplete_Click()
  ADOprimaryrs![PROJ Saved] = True  '
  ADOprimaryrs![PROJ Custom 1] = FormatDate(Now) & ""
  ADOprimaryrs.Update
  Label2.Caption = "Closed On " & FormatDate(ADOprimaryrs![PROJ Custom 1])
  EnabledSub False
End Sub

Private Sub cmdDate_Click(Index As Integer)
Select Case Index
Case 4
    Menu_Calendar.WhoCallMe True, 1600
Case 12
    Menu_Calendar.WhoCallMe True, 1610
Case 0
    Menu_Calendar.WhoCallMe True, 1645  '17
Case 1
    Menu_Calendar.WhoCallMe True, 1650  '16
End Select
    'Menu_Calendar.Show vbModal
End Sub

Private Sub cmdProjectID_Click()
    Dim ghead As String
    Dim fhead As String
    
    ghead = "Projects"
    fhead = "ID//Name//Manager"
    AllLookup.ToWhichRecord ADOprimaryrs, ghead, fhead
    'AllLookup.Show vbModal
End Sub

Private Sub cmdSearch_Click()
If ADOTransRS Is Nothing Then
Else
    If ADOTransRS.RecordCount = 0 Then Exit Sub
    SearchRECORD ADOTransRS, grddatagrid, txtfields(14).Text, lblLabels(17).Caption, WhichField, "AR SALE Ext Document #"
End If
End Sub

Private Sub cmdShow_Click()
    If ADOTransRS Is Nothing Then
    Else
        ADOTransRS.Close
        Set ADOTransRS = Nothing
    End If
    
    Set ADOTransRS = New ADODB.Recordset

    TempStr = "SELECT DISTINCTROW [AR Sales].[AR SALE Posted YN],[AR Sales].[AR SALE Ext Document #]," & _
    "[AR Sales].[AR SALE Salesperson],[AR Sales].[AR SALE PO ID],[AR Sales].[AR SALE Document Type]," & _
    "[AR Sales].[AR SALE Customer ID], [AR Sales].[AR SALE Date],[AR Sales].[AR SALE Total] " & _
    "FROM [AR Sales] INNER JOIN [AR Sales Detail] ON [AR Sales].[AR SALE Document #] = " & _
    "[AR Sales Detail].[AR SALED Document #] WHERE [AR Sales].[AR SALE Date] BETWEEN #" & txtfields(17).Text & "# AND #" & txtfields(16).Text & "#"
    'Debug.Print TempStr
     ADOTransRS.Open TempStr, db, adOpenStatic, adLockReadOnly, adCmdText
  'End If
     Set grddatagrid.DataSource = ADOTransRS
     If ADOTransRS.RecordCount = 0 Then
        MsgBox "There is no Project transaction"
        Exit Sub
     End If
End Sub

Private Sub Form_Load()
  Me.Width = 9900
  Me.Height = 7455

On Error GoTo FormErr
ShowStatus True
    Set db = New ADODB.Connection
    db.CursorLocation = adUseClient
    db.Open gblADOProvider

  Dim sql As String
  sql = "select [PROJ ID],[PROJ Name],[PROJ Project Manager],[PROJ Start Date]," & _
    "[PROJ Completion Date],[PROJ Type],[PROJ Description],[PROJ Notes]," & _
    "[PROJ Estimated Cost],[PROJ Actual Cost],[PROJ Estimated Revenue]," & _
    "[PROJ Actual Revenue],[PROJ Customer ID],[PROJ Saved],[PROJ Custom 1] from [PROJ Projects] " & _
    "Order by [PROJ ID]"
  
  Set ADOprimaryrs = New ADODB.Recordset
  ADOprimaryrs.Open sql, db, adOpenStatic, adLockOptimistic
  
  'Lock these to prevent invalid entries & unwanted modifications
  txtfields(0).Locked = True
  txtfields(6).Locked = True
  txtfields(10).Locked = True
  txtfields(11).Locked = True
  Dim i As Integer
'  For i = 13 To 19
'    txtfields(i).Locked = True
'  Next i
  
  Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtfields
    Set oText.DataSource = ADOprimaryrs
    If oText.DataField <> "" Then
        If ADOprimaryrs("" & oText.DataField & "").Type = 202 Then oText.MaxLength = ADOprimaryrs("" & oText.DataField & "").DefinedSize
    End If
  Next
  
  Set Me.cbfields.DataSource = ADOprimaryrs
  
    loadCombo
    
  If CheckNewDB(ADOprimaryrs, "Projects") = True Then
    cmdAdd_Click
  Else
    RefreshUnboundText
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
  
  Me.Width = 9900
  Me.Height = 7455
  
SkipResize:
  frPrimary.Left = (Me.ScaleWidth - frPrimary.Width) / 2
  Label1.Left = frPrimary.Left
  Label1.Width = frPrimary.Width
  frPrimary.Top = (Me.ScaleHeight - frPrimary.Height - picButtons.Height - Me.picStatBox.Height) / 2 + 230
  AR_Cust_Proj_Drill.Top = frPrimary.Top
  AR_Cust_Proj_Drill.Left = frPrimary.Left
  
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
    'updates the checklist Projects
  ShowStatus True
      EndLoad db, ADOprimaryrs, "Projects"
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
  Set frm_AR_Cust_Projects = Nothing
  Exit Sub
FormErr:
  MsgBox Err.Description
  ShowStatus False
End Sub

Private Sub adoPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
ShowStatus True
Dim EnabledButt As Boolean
  If Not ADOprimaryrs.EOF And Not ADOprimaryrs.BOF Then
    If mbAddNewFlag = False Then RefreshUnboundText
  Else
    ShowStatus False
    Exit Sub
  End If
  If ADOprimaryrs![PROJ Saved] = True Then 'PROJ Custom 1
    EnabledSub False
    Label2.Caption = "Closed On " & FormatDate(ADOprimaryrs![PROJ Custom 1])
  Else
    EnabledSub True
  End If
      
  'This will display the current record position for this recordset
  lblStatus.Caption = "Record: " & CStr(ADOprimaryrs.AbsolutePosition) & " of " & CStr(ADOprimaryrs.RecordCount)
ShowStatus False
End Sub

Private Sub EnabledSub(EnabledButt As Boolean)
    Image1.Visible = Not EnabledButt
    Label2.Visible = Not EnabledButt
    picMajor.Enabled = EnabledButt
    cmdComplete.Enabled = EnabledButt
    ReCalcButton.Enabled = EnabledButt
    cmdDelete.Enabled = EnabledButt
    cmdUpdate.Enabled = EnabledButt
    GetTextColor Me
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
        If Not (.BOF And .EOF) Then
          mvBookMark = .Bookmark
        End If
        mbAddNewFlag = True
        .AddNew
        lblStatus.Caption = "Add record"
        txtfields(0).Locked = False
        lblfields(0) = ""
        lblfields(1) = ""
        cmdAdd.Caption = "&Cancel"
        btCust.Enabled = True
        cmdUpdate.Enabled = True
        
        Dim oText As TextBox
        'clear unbound textbox
        For Each oText In Me.txtfields
          If oText.DataField = "" Then
              oText.Text = ""
          End If
        Next
    Else
        mbAddNewFlag = False
        .CancelUpdate
        txtfields(0).Locked = True
        If .RecordCount > 0 Then
            If mvBookMark > 0 Then
                .Bookmark = mvBookMark
            Else
                .MoveFirst
            End If
        Else
            btCust.Enabled = False
        End If
        cmdAdd.Caption = "&Add"
    End If
  End With
  
    'set the controls accordingly
    cmdDelete.Enabled = Not mbAddNewFlag
    cmdRefresh.Enabled = Not mbAddNewFlag
    ReCalcButton.Enabled = Not mbAddNewFlag
    TransButton.Enabled = Not mbAddNewFlag
    cmdComplete.Enabled = Not mbAddNewFlag
  GetTextColor Me
    
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
  'On Error GoTo DeleteErr
  With ADOprimaryrs
    If .RecordCount = 0 Then Exit Sub   ' no records maa....
    If .EditMode = False Then
        .Delete
        .MoveNext
        If .RecordCount = 0 Then  ' no more records
            cmdUpdate.Enabled = False
            cmdDelete.Enabled = False
            cmdRefresh.Enabled = False
            ReCalcButton.Enabled = False
            TransButton.Enabled = False
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
            Dim srch As String
            srch = txtfields(0).Text
            .Requery
            .Find "[PROJ ID] = '" & srch & "'"
        End If
    End With
    RefreshUnboundText
    
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub

Private Sub cmdUpdate_Click()
On Error GoTo UpdateErr

    With ADOprimaryrs
        If .RecordCount = 0 Then Exit Sub 'no records to update
        If Trim(txtfields(0).Text) <> "" Then
        Dim oTxt As TextBox
          For Each oTxt In Me.txtfields
            If oTxt.Text = "" Then
                If oTxt.DataField <> "" Then
                  If ADOprimaryrs("" & oTxt.DataField & "").Type = 203 Or ADOprimaryrs("" & oTxt.DataField & "").Type = 202 Then oTxt.Text = " "
                End If
            End If
          Next
        Else
            MsgBox lblLabels(0) & " must be filled. Please try again before Update.", vbInformation, "Information"
            Exit Sub
        End If
        .Update
        Dim srch As String
        srch = txtfields(0).Text
        If mbAddNewFlag Then 'requery to get default value assigned by database
            .Requery
            .Find "[PROJ ID] = '" & srch & "'" ' go to newly created record
            mbAddNewFlag = False
        End If
        
        'reenable the necessary buttons
        cmdAdd.Caption = "&Add"
        txtfields(0).Locked = True
        cmdDelete.Enabled = True
        cmdRefresh.Enabled = True
        cmdProjectID.Enabled = True
        ReCalcButton.Enabled = True
        TransButton.Enabled = True
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
    RefreshUnboundText
  Exit Sub

GoFirstError:
  MsgBox Err.Description
End Sub

Private Sub cmdLast_Click()
  'On Error GoTo GoLastError
    ADOprimaryrs.MoveLast
    RefreshUnboundText

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
    RefreshUnboundText
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
    RefreshUnboundText
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
If ADOTransRS.RecordCount = 0 Then Exit Sub
    lblLabels(17) = grddatagrid.Columns(ColIndex).Caption
    WhichField = grddatagrid.Columns(ColIndex).DataField
    ADOTransRS.Close
    Set ADOTransRS = Nothing
    Set ADOTransRS = New ADODB.Recordset
    ADOTransRS.Open TempStr & " ORDER BY [" & WhichField & "]", db, adOpenKeyset, adLockOptimistic, adCmdText
    Set grddatagrid.DataSource = ADOTransRS
End Sub

Private Sub ReCalcButton_Click()
  'On Error GoTo RecalcFinancials_Error

  'This procedure calculates the actual cost of the project from the sales, inventory, purchases and project history tables
  
  Dim ActualCost As Double
  Dim ActualRevenue As Double
  Dim ProjectID As String
  Dim qryPurchases As String
  Dim qryARSales As String
  Dim qryInventoryAdjustment As String
    
  'On Error Resume Next

  'Make sure the Porject ID is Valid Before Continuing
  If IsNull(Trim(Me.txtfields(0).Text)) Then
    ActualCost = 0
    ActualRevenue = 0
    GoTo SkipIt
  End If
  
  'Assign Variables
  ProjectID = Me.txtfields(0).Text
  qryPurchases = " [AP Purchase] INNER JOIN [AP Purchase Detail] ON [AP Purchase].[AP PO Document No] = [AP Purchase Detail].[AP POD Document No]"
  qryARSales = " [AR Sales] INNER JOIN [AR Sales Detail] ON [AR Sales].[AR SALE Document #] = [AR Sales Detail].[AR SALED Document #]"
  qryInventoryAdjustment = " [INV Adjustment] INNER JOIN [INV Adjustment Detail] ON [INV Adjustment].[INV ADJ Document No] = [INV Adjustment Detail].[INV ADJD Document No]"
  
  'Calculate Actual Cost
  ActualCostA = SumRecord("[AP POD Item Total]", qryPurchases, db, "[AP POD Project ID] = '" & ProjectID & "' AND [AP PO Posted YN] = TRUE AND [AP PO Document Type] in ('Receiving','Voucher')")
  If IsNull(ActualCostA) Then ActualCostA = 0
  
  ActualCostB = SumRecord("[PROJ HIST Cost]", "[PROJ History]", db, "[PROJ HIST Project ID] = '" & ProjectID & "'")
  If IsNull(ActualCostB) Then ActualCostB = 0
  
  ActualCostC = SumRecord("([INV ADJD Cost] * [INV ADJD Adjusted Qty])", qryInventoryAdjustment, db, "[INV ADJD Project] = '" & ProjectID & "' AND [INV ADJ Type] = 'Decrease' AND [INV ADJ Posted YN] = TRUE")
  If IsNull(ActualCostC) Then ActualCostC = 0
  
  ActualCostD = SumRecord("[AP POD Item Total]", qryPurchases, db, "[AP POD Project ID] = '" & ProjectID & "' AND [AP PO Posted YN] = TRUE AND [AP PO Document Type] in ('Credit Memo')")
  If IsNull(ActualCostD) Then ActualCostD = 0
  
  ActualCost = (ActualCostA + ActualCostB + ActualCostC) - ActualCostD

  'Calculate Actual Revenue
  ActualRevenueA = SumRecord("[AR SALED Item Total]", qryARSales, db, "[AR SALED Project] = '" & ProjectID & "' AND [AR SALE Posted YN] = TRUE AND [AR SALE Document Type] in ('Invoice','Sales Memo')")
  If IsNull(ActualRevenueA) Then ActualRevenueA = 0
  
  ActualRevenueB = SumRecord("[PROJ HIST Revenue]", "[PROJ History]", db, "[PROJ HIST Project ID] = '" & ProjectID & "'")
  If IsNull(ActualRevenueB) Then ActualRevenueB = 0
  
  ActualRevenueC = SumRecord("[INV ADJD Cost] * [INV ADJD Adjusted Qty]", qryInventoryAdjustment, db, "[INV ADJD Project] = '" & ProjectID & "' AND [INV ADJ Type] = 'Increase' AND [INV ADJ Posted YN] = TRUE")
  If IsNull(ActualRevenueC) Then ActualRevenueC = 0
  
  ActualRevenueD = SumRecord("[AR SALED Item Total]", qryARSales, db, "[AR SALED Project] = '" & ProjectID & "' AND [AR SALE Posted YN] = TRUE AND [AR SALE Document Type] in ('Return','Credit Memo')")
  If IsNull(ActualRevenueD) Then ActualRevenueD = 0
  
  ActualRevenue = (ActualRevenueA + ActualRevenueB + ActualRevenueC) - ActualRevenueD

SkipIt:
  ' Assign Values to Approiate Fields
  ADOprimaryrs.Fields("PROJ Actual Revenue") = ActualRevenue
  ADOprimaryrs.Fields("PROJ Actual Cost") = ActualCost

  'Refersh the record set and display the new results
  ADOprimaryrs.UpdateBatch adAffectAll
  ADOprimaryrs.Requery
  
    RefreshUnboundText
  MsgBox "Calculation completed"
  Exit Sub
RecalcFinancials_Error:
  Resume Next

End Sub

Private Sub TransButton_Click()
If txtfields(0).Text = "" And txtfields(6).Text = "" Then
    MsgBox "Please select Project ID"
    Exit Sub
End If
Set ADOTransRS = New ADODB.Recordset

    TempStr = "SELECT DISTINCTROW [AR Sales].[AR SALE Posted YN],[AR Sales].[AR SALE Ext Document #]," & _
    "[AR Sales].[AR SALE Salesperson],[AR Sales].[AR SALE PO ID],[AR Sales].[AR SALE Document Type]," & _
    "[AR Sales].[AR SALE Customer ID], [AR Sales].[AR SALE Date],[AR Sales].[AR SALE Total] " & _
    "FROM [AR Sales] INNER JOIN [AR Sales Detail] ON [AR Sales].[AR SALE Document #] = " & _
    "[AR Sales Detail].[AR SALED Document #]  WHERE [AR SALED Project] = '" & txtfields(0).Text & "'"
'    ADOprimaryrs.Open "SELECT DISTINCTROW [AR Sales].[AR SALE Document #], [AR Sales].[AR SALE Date], [AR Sales].[AR SALE Document Type], [AR Sales].[AR SALE Total], [AR Sales Detail].[AR SALED Project] FROM [AR Sales] INNER JOIN [AR Sales Detail] ON [AR Sales].[AR SALE Document #] = [AR Sales Detail].[AR SALED Document #]", db, adOpenStatic, adLockOptimistic
'  Else
     ADOTransRS.Open TempStr, db, adOpenStatic, adLockReadOnly, adCmdText
  'End If
     Set grddatagrid.DataSource = ADOTransRS
     If ADOTransRS.RecordCount = 0 Then
        MsgBox "There is no Project transaction"
        Exit Sub
     End If
    AR_Cust_Proj_Drill.ZOrder 0
    AR_Cust_Proj_Drill.Visible = True
    Form_Resize
    cmdAdd.Enabled = False
    cmdUpdate.Enabled = False
    cmdDelete.Enabled = False
    cmdRefresh.Enabled = False

End Sub

Private Sub loadCombo()
    ComboInit cbfields, lblLabels(5), "select [LIST PROJECT Types] as Projects from " & _
        "[LIST Project Types]"
End Sub

Private Sub RefreshUnboundText()
'this sub refreshes the textboxes that are not bound to the recordset
    If txtfields(6).Text <> "" Then
        lblfields(0) = LookRecord("[AR CUST Name]", "[AR Customer]", db, "[AR CUST Customer ID] = '" & Me.txtfields(6).Text & "'")
        lblfields(1) = LookRecord("[AR CUST Phone]", "[AR Customer]", db, "[AR CUST Customer ID] = '" & Me.txtfields(6).Text & "'") & " Ext. " & LookRecord("[AR CUST Phone Ext]", "[AR Customer]", db, "[AR CUST Customer ID] = '" & Me.txtfields(6).Text & "'")
    Else
        lblfields(0) = ""
        lblfields(1) = ""
    End If
    Me.txtfields(15).Text = FormatCurr(ADOprimaryrs.Fields("PROJ Estimated Cost") - ADOprimaryrs.Fields("PROJ Actual Cost"))
    Me.txtfields(21).Text = FormatCurr(ADOprimaryrs.Fields("PROJ Estimated Revenue") - ADOprimaryrs.Fields("PROJ Actual Revenue"))
    Me.txtfields(20).Text = FormatCurr(ADOprimaryrs.Fields("PROJ Estimated Revenue") - ADOprimaryrs.Fields("PROJ Estimated Cost"))
    Me.txtfields(18).Text = FormatCurr(ADOprimaryrs.Fields("PROJ Actual Revenue") - ADOprimaryrs.Fields("PROJ Actual Cost"))
    Me.txtfields(19).Text = FormatCurr(Me.txtfields(20).Text - Me.txtfields(18).Text)
End Sub
