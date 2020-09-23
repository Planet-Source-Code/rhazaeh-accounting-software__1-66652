VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Begin VB.Form frm_SYS_Setup_Payroll 
   Caption         =   "Payroll Setup"
   ClientHeight    =   7350
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   9450
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7350
   ScaleWidth      =   9450
   Begin VB.PictureBox frprimary 
      BorderStyle     =   0  'None
      Height          =   6975
      Left            =   0
      ScaleHeight     =   6975
      ScaleWidth      =   9375
      TabIndex        =   4
      Top             =   0
      Width           =   9375
      Begin VB.PictureBox picOptions 
         BorderStyle     =   0  'None
         Height          =   6060
         Index           =   0
         Left            =   360
         ScaleHeight     =   6157.742
         ScaleMode       =   0  'User
         ScaleWidth      =   8898.92
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   840
         Width           =   8805
         Begin VB.CommandButton cmdAcct 
            Height          =   255
            Index           =   0
            Left            =   3480
            Picture         =   "frm_SYS_Setup_Payroll.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   88
            Top             =   1440
            Width           =   375
         End
         Begin VB.TextBox txtFields 
            DataField       =   "FIT PAY ACCT"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   0
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   87
            Text            =   "231100"
            Top             =   1440
            Width           =   1935
         End
         Begin VB.TextBox txtFields 
            DataField       =   "FICA PAY ACCT"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   1
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   86
            Text            =   "231200"
            Top             =   2160
            Width           =   1935
         End
         Begin VB.CommandButton cmdAcct 
            Height          =   255
            Index           =   1
            Left            =   3480
            Picture         =   "frm_SYS_Setup_Payroll.frx":014A
            Style           =   1  'Graphical
            TabIndex        =   85
            Top             =   2160
            Width           =   375
         End
         Begin VB.TextBox txtFields 
            DataField       =   "FUTA PAY ACCT"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   2
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   84
            Text            =   "231300"
            Top             =   2880
            Width           =   1935
         End
         Begin VB.CommandButton cmdAcct 
            Height          =   255
            Index           =   2
            Left            =   3480
            Picture         =   "frm_SYS_Setup_Payroll.frx":0294
            Style           =   1  'Graphical
            TabIndex        =   83
            Top             =   2880
            Width           =   375
         End
         Begin VB.TextBox txtFields 
            DataField       =   "WAGES PAY ACCT"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   3
            Left            =   5760
            Locked          =   -1  'True
            TabIndex        =   82
            Text            =   "241000"
            Top             =   1440
            Width           =   1935
         End
         Begin VB.CommandButton cmdAcct 
            Height          =   255
            Index           =   3
            Left            =   7680
            Picture         =   "frm_SYS_Setup_Payroll.frx":03DE
            Style           =   1  'Graphical
            TabIndex        =   81
            Top             =   1440
            Width           =   375
         End
         Begin VB.TextBox txtFields 
            DataField       =   "SUI PAY ACCT"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   4
            Left            =   5760
            Locked          =   -1  'True
            TabIndex        =   80
            Text            =   "231400"
            Top             =   2160
            Width           =   1935
         End
         Begin VB.CommandButton cmdAcct 
            Height          =   255
            Index           =   4
            Left            =   7680
            Picture         =   "frm_SYS_Setup_Payroll.frx":0528
            Style           =   1  'Graphical
            TabIndex        =   79
            Top             =   2160
            Width           =   375
         End
         Begin VB.TextBox txtFieldsTemp 
            DataField       =   " "
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Left            =   5520
            Locked          =   -1  'True
            TabIndex        =   78
            Text            =   "231100"
            Top             =   5040
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.Label lblLabels 
            Caption         =   "FIT Payable:"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   100
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label lblAcct 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   99
            Top             =   1725
            Width           =   3495
         End
         Begin VB.Label lblLabels 
            Caption         =   "FICA Payable:"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   98
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label lblAcct 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   97
            Top             =   2445
            Width           =   3495
         End
         Begin VB.Label lblLabels 
            Caption         =   "FUTA Payable:"
            Height          =   255
            Index           =   3
            Left            =   360
            TabIndex        =   96
            Top             =   2880
            Width           =   1215
         End
         Begin VB.Label lblAcct 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   95
            Top             =   3165
            Width           =   3495
         End
         Begin VB.Label lblLabels 
            Caption         =   "Wages Payable:"
            Height          =   255
            Index           =   4
            Left            =   4560
            TabIndex        =   94
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label lblAcct 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   4560
            TabIndex        =   93
            Top             =   1725
            Width           =   3495
         End
         Begin VB.Label lblLabels 
            Caption         =   "SUI Payable:"
            Height          =   255
            Index           =   5
            Left            =   4560
            TabIndex        =   92
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label lblAcct 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   4
            Left            =   4560
            TabIndex        =   91
            Top             =   2445
            Width           =   3495
         End
         Begin VB.Label lblAcctTemp 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   5520
            TabIndex        =   90
            Top             =   4800
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000001&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Liability Accounts"
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   89
            Top             =   0
            Width           =   8775
         End
      End
      Begin ComctlLib.TabStrip tbPay 
         Height          =   6615
         Left            =   240
         TabIndex        =   76
         Top             =   360
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   11668
         _Version        =   327682
         BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
            NumTabs         =   3
            BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Liability Accounts"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Expense Account"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Rates/Others"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin VB.PictureBox picOptions 
         BorderStyle     =   0  'None
         Height          =   6060
         Index           =   2
         Left            =   360
         ScaleHeight     =   6157.742
         ScaleMode       =   0  'User
         ScaleWidth      =   8898.92
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   840
         Width           =   8805
         Begin VB.Frame Frame2 
            Caption         =   "Employer/Employee Faderal Taxes"
            Height          =   3975
            Left            =   0
            TabIndex        =   56
            Top             =   360
            Width           =   4335
            Begin VB.TextBox txtRO 
               DataField       =   "FICASS"
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
               Index           =   13
               Left            =   2520
               TabIndex        =   63
               Top             =   600
               Width           =   855
            End
            Begin VB.TextBox txtRO 
               DataField       =   "FICAMED"
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
               Index           =   12
               Left            =   2520
               TabIndex        =   62
               Top             =   1080
               Width           =   855
            End
            Begin VB.TextBox txtRO 
               DataField       =   "FICA EMPL PERCENT"
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
               Index           =   11
               Left            =   2520
               TabIndex        =   61
               Top             =   1560
               Width           =   855
            End
            Begin VB.TextBox txtRO 
               DataField       =   "MEDIWAGEBASE"
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
               Left            =   2520
               TabIndex        =   60
               Top             =   2040
               Width           =   1335
            End
            Begin VB.TextBox txtRO 
               DataField       =   "SSWAGEBASE"
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
               Left            =   2520
               TabIndex        =   59
               Top             =   2520
               Width           =   1335
            End
            Begin VB.TextBox txtRO 
               DataField       =   "FUTARATE"
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
               Index           =   8
               Left            =   2520
               TabIndex        =   58
               Top             =   3000
               Width           =   855
            End
            Begin VB.TextBox txtRO 
               DataField       =   "FUTAWAGEBASE"
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
               Left            =   2520
               TabIndex        =   57
               Top             =   3480
               Width           =   1335
            End
            Begin VB.Label lblRO 
               Caption         =   "FICA Social Sec Rate:"
               Height          =   255
               Index           =   13
               Left            =   480
               TabIndex        =   74
               Top             =   600
               Width           =   1815
            End
            Begin VB.Label lblRO 
               Caption         =   "FICA Medicare Rate:"
               Height          =   255
               Index           =   12
               Left            =   480
               TabIndex        =   73
               Top             =   1080
               Width           =   1815
            End
            Begin VB.Label lblRO 
               Caption         =   "FICA Employer Portion:"
               Height          =   255
               Index           =   11
               Left            =   480
               TabIndex        =   72
               Top             =   1560
               Width           =   1815
            End
            Begin VB.Label lblRO 
               Caption         =   "Medicare Wage Base:"
               Height          =   255
               Index           =   10
               Left            =   480
               TabIndex        =   71
               Top             =   2040
               Width           =   1815
            End
            Begin VB.Label lblRO 
               Caption         =   "Social Sec. Wages Base:"
               Height          =   255
               Index           =   9
               Left            =   480
               TabIndex        =   70
               Top             =   2520
               Width           =   1815
            End
            Begin VB.Label lblRO 
               Caption         =   "FUTA Rate:"
               Height          =   255
               Index           =   8
               Left            =   480
               TabIndex        =   69
               Top             =   3000
               Width           =   1815
            End
            Begin VB.Label lblRO 
               Caption         =   "FUTA Wages Base:"
               Height          =   255
               Index           =   7
               Left            =   480
               TabIndex        =   68
               Top             =   3480
               Width           =   1815
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "%"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   7
               Left            =   3480
               TabIndex        =   67
               Top             =   600
               Width           =   210
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "%"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   6
               Left            =   3480
               TabIndex        =   66
               Top             =   1080
               Width           =   210
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "%"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   5
               Left            =   3480
               TabIndex        =   65
               Top             =   1560
               Width           =   210
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "%"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   4
               Left            =   3480
               TabIndex        =   64
               Top             =   3000
               Width           =   210
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Employer State Taxes"
            Height          =   1575
            Left            =   0
            TabIndex        =   50
            Top             =   4440
            Width           =   4335
            Begin VB.TextBox txtRO 
               DataField       =   "SUIRATE"
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
               Index           =   0
               Left            =   2520
               TabIndex        =   52
               Top             =   600
               Width           =   855
            End
            Begin VB.TextBox txtRO 
               DataField       =   "SUIWAGEBASE"
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
               Index           =   1
               Left            =   2520
               TabIndex        =   51
               Top             =   1080
               Width           =   1335
            End
            Begin VB.Label lblRO 
               Caption         =   "SUI Rate:"
               Height          =   255
               Index           =   0
               Left            =   480
               TabIndex        =   55
               Top             =   600
               Width           =   1815
            End
            Begin VB.Label lblRO 
               Caption         =   "SUI Wages Base:"
               Height          =   255
               Index           =   1
               Left            =   480
               TabIndex        =   54
               Top             =   1080
               Width           =   1815
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "%"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   0
               Left            =   3480
               TabIndex        =   53
               Top             =   600
               Width           =   210
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "W2 Fields"
            Height          =   3015
            Left            =   4440
            TabIndex        =   41
            Top             =   360
            Width           =   4335
            Begin VB.TextBox txtRO 
               DataField       =   "FederalIDNo"
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
               Index           =   2
               Left            =   2520
               TabIndex        =   45
               Top             =   720
               Width           =   1335
            End
            Begin VB.TextBox txtRO 
               DataField       =   "StateIDNo"
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
               Index           =   3
               Left            =   2520
               TabIndex        =   44
               Top             =   1200
               Width           =   1335
            End
            Begin VB.TextBox txtRO 
               DataField       =   "State"
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
               Index           =   4
               Left            =   2520
               TabIndex        =   43
               Top             =   1680
               Width           =   1335
            End
            Begin VB.TextBox txtRO 
               DataField       =   "LocalityName"
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
               Index           =   5
               Left            =   2520
               TabIndex        =   42
               Top             =   2160
               Width           =   1335
            End
            Begin VB.Label lblRO 
               Caption         =   "Employer's Fed ID No:"
               Height          =   255
               Index           =   2
               Left            =   480
               TabIndex        =   49
               Top             =   720
               Width           =   1815
            End
            Begin VB.Label lblRO 
               Caption         =   "Employer's State ID No:"
               Height          =   255
               Index           =   3
               Left            =   480
               TabIndex        =   48
               Top             =   1200
               Width           =   1815
            End
            Begin VB.Label lblRO 
               Caption         =   "State Abbreviation:"
               Height          =   255
               Index           =   4
               Left            =   480
               TabIndex        =   47
               Top             =   1680
               Width           =   1815
            End
            Begin VB.Label lblRO 
               Caption         =   "Locality Name:"
               Height          =   255
               Index           =   5
               Left            =   480
               TabIndex        =   46
               Top             =   2160
               Width           =   1815
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Employee Earning"
            Height          =   2535
            Left            =   4440
            TabIndex        =   36
            Top             =   3480
            Width           =   4335
            Begin VB.TextBox txtRO 
               DataField       =   "OTAfter"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
               Height          =   285
               Index           =   6
               Left            =   1920
               TabIndex        =   38
               Top             =   720
               Width           =   495
            End
            Begin VB.CheckBox chkAuto 
               Caption         =   "Automatically Pay All Commissionable Invoice Less Return"
               DataField       =   "PAYALLCOMMISSIONS"
               Height          =   375
               Left            =   480
               TabIndex        =   37
               Top             =   1440
               Width           =   3375
            End
            Begin VB.Label lblRO 
               Caption         =   "Overtime Paid After"
               Height          =   255
               Index           =   6
               Left            =   480
               TabIndex        =   40
               Top             =   720
               Width           =   1455
            End
            Begin VB.Label lblRO 
               Caption         =   "Hours Per Week"
               Height          =   255
               Index           =   14
               Left            =   2520
               TabIndex        =   39
               Top             =   720
               Width           =   1455
            End
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000001&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Rates/Others"
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   75
            Top             =   0
            Width           =   8775
         End
      End
      Begin VB.PictureBox picOptions 
         BorderStyle     =   0  'None
         Height          =   6060
         Index           =   1
         Left            =   360
         ScaleHeight     =   6157.742
         ScaleMode       =   0  'User
         ScaleWidth      =   8868.601
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   840
         Width           =   8775
         Begin VB.TextBox txtFields 
            DataField       =   "FICA EXP ACCT"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   5
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   19
            Text            =   "725100"
            Top             =   1440
            Width           =   1935
         End
         Begin VB.CommandButton cmdAcct 
            Height          =   255
            Index           =   5
            Left            =   3720
            Picture         =   "frm_SYS_Setup_Payroll.frx":0672
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   1440
            Width           =   375
         End
         Begin VB.TextBox txtFields 
            DataField       =   "FUTA EXP ACCT"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   6
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   17
            Text            =   "725200"
            Top             =   2160
            Width           =   1935
         End
         Begin VB.CommandButton cmdAcct 
            Height          =   255
            Index           =   6
            Left            =   3720
            Picture         =   "frm_SYS_Setup_Payroll.frx":07BC
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   2160
            Width           =   375
         End
         Begin VB.TextBox txtFields 
            DataField       =   "SUI EXP ACCT"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   7
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   15
            Text            =   "725300"
            Top             =   2880
            Width           =   1935
         End
         Begin VB.CommandButton cmdAcct 
            Height          =   255
            Index           =   7
            Left            =   3720
            Picture         =   "frm_SYS_Setup_Payroll.frx":0906
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   2880
            Width           =   375
         End
         Begin VB.TextBox txtFields 
            DataField       =   "SALES EXP ACCT"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   8
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   13
            Text            =   "720100"
            Top             =   3600
            Width           =   1935
         End
         Begin VB.CommandButton cmdAcct 
            Height          =   255
            Index           =   8
            Left            =   3720
            Picture         =   "frm_SYS_Setup_Payroll.frx":0A50
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   3600
            Width           =   375
         End
         Begin VB.TextBox txtFields 
            DataField       =   "OFFICE EXP ACCT"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   9
            Left            =   6360
            Locked          =   -1  'True
            TabIndex        =   11
            Text            =   "720200"
            Top             =   1440
            Width           =   1935
         End
         Begin VB.CommandButton cmdAcct 
            Height          =   255
            Index           =   9
            Left            =   8280
            Picture         =   "frm_SYS_Setup_Payroll.frx":0B9A
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   1440
            Width           =   375
         End
         Begin VB.TextBox txtFields 
            DataField       =   "WHSE EXP ACCT"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   10
            Left            =   6360
            Locked          =   -1  'True
            TabIndex        =   9
            Text            =   "720300"
            Top             =   2160
            Width           =   1935
         End
         Begin VB.CommandButton cmdAcct 
            Height          =   255
            Index           =   10
            Left            =   8280
            Picture         =   "frm_SYS_Setup_Payroll.frx":0CE4
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   2160
            Width           =   375
         End
         Begin VB.TextBox txtFields 
            DataField       =   "PROD EXP ACCT"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   11
            Left            =   6360
            Locked          =   -1  'True
            TabIndex        =   7
            Text            =   "720400"
            Top             =   2880
            Width           =   1935
         End
         Begin VB.CommandButton cmdAcct 
            Height          =   255
            Index           =   11
            Left            =   8280
            Picture         =   "frm_SYS_Setup_Payroll.frx":0E2E
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   2880
            Width           =   375
         End
         Begin VB.Label lblAcct 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   5
            Left            =   360
            TabIndex        =   34
            Top             =   1725
            Width           =   3735
         End
         Begin VB.Label lblLabels 
            Caption         =   "FICA Exp:"
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   33
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label lblAcct 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   6
            Left            =   360
            TabIndex        =   32
            Top             =   2445
            Width           =   3735
         End
         Begin VB.Label lblLabels 
            Caption         =   "FUTA Exp:"
            Height          =   255
            Index           =   6
            Left            =   360
            TabIndex        =   31
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label lblAcct 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   7
            Left            =   360
            TabIndex        =   30
            Top             =   3165
            Width           =   3735
         End
         Begin VB.Label lblLabels 
            Caption         =   "SUI Exp:"
            Height          =   255
            Index           =   7
            Left            =   360
            TabIndex        =   29
            Top             =   2880
            Width           =   1215
         End
         Begin VB.Label lblAcct 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   8
            Left            =   360
            TabIndex        =   28
            Top             =   3885
            Width           =   3735
         End
         Begin VB.Label lblLabels 
            Caption         =   "Sales Salaries Exp:"
            Height          =   255
            Index           =   8
            Left            =   360
            TabIndex        =   27
            Top             =   3600
            Width           =   1455
         End
         Begin VB.Label lblAcct 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   9
            Left            =   4440
            TabIndex        =   26
            Top             =   1725
            Width           =   4215
         End
         Begin VB.Label lblLabels 
            Caption         =   "Office Salaries Exp:"
            Height          =   255
            Index           =   9
            Left            =   4440
            TabIndex        =   25
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label lblAcct 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   10
            Left            =   4440
            TabIndex        =   24
            Top             =   2445
            Width           =   4215
         End
         Begin VB.Label lblLabels 
            Caption         =   "Warehouse Salaries Exp:"
            Height          =   255
            Index           =   10
            Left            =   4440
            TabIndex        =   23
            Top             =   2160
            Width           =   1815
         End
         Begin VB.Label lblAcct 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   11
            Left            =   4440
            TabIndex        =   22
            Top             =   3165
            Width           =   4215
         End
         Begin VB.Label lblLabels 
            Caption         =   "Production Salaries Exp:"
            Height          =   255
            Index           =   11
            Left            =   4440
            TabIndex        =   21
            Top             =   2880
            Width           =   1815
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000001&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Expense Accounts"
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   20
            Top             =   0
            Width           =   8775
         End
      End
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   9450
      TabIndex        =   0
      Top             =   7050
      Width           =   9450
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   300
         Left            =   2400
         TabIndex        =   2
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   300
         Left            =   1320
         TabIndex        =   1
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   300
         Left            =   240
         TabIndex        =   3
         Top             =   0
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frm_SYS_Setup_Payroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ADOprimaryrs As ADODB.Recordset
Attribute ADOprimaryrs.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean

Dim db As ADODB.Connection

Private Sub chkAuto_Click()
If chkAuto.Value = 0 Then _
    db.Execute "Update [AR SALES] SET [AR SALE Select to Pay]= 0 WHERE [AR SALE Commission Paid] = 0"
End Sub

'The recordset should only contain one record holding information pertaining to  a specific company
' inventory setup.

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
    lblAcct(Index).Caption = lblAcctTemp.Caption
    txtfields(Index).Text = txtFieldsTemp.Text
    txtfields(Index).SetFocus
End Sub

Private Sub Form_Load()
On Error GoTo FormErr
ShowStatus True
  Set db = New ADODB.Connection
  db.CursorLocation = adUseClient
  db.Open gblADOProvider

  Set ADOprimaryrs = New ADODB.Recordset
  ADOprimaryrs.Open "Select * from [Pyrl - Setup]", db, adOpenKeyset, adLockOptimistic, adCmdText
  Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtRO
    Set oText.DataSource = ADOprimaryrs
    If ADOprimaryrs("" & oText.DataField & "").Type = 202 Then oText.MaxLength = ADOprimaryrs("" & oText.DataField & "").DefinedSize
  Next
  
  Set chkAuto.DataSource = ADOprimaryrs
  'Bind the Check boxes to the data provider
  For Each oText In Me.txtfields
    Set oText.DataSource = ADOprimaryrs
    If ADOprimaryrs("" & oText.DataField & "").Type = 202 Then oText.MaxLength = ADOprimaryrs("" & oText.DataField & "").DefinedSize
    If ADOprimaryrs.RecordCount > 0 And oText.Text <> "" Then
        lblAcct(oText.Index).Caption = LookRecord("[GL COA Account Name]", "[GL Chart Of Accounts]", db, "[GL COA Account No] = '" & oText.Text & "'")
    End If
  Next
  picOptions(0).Left = 360
  picOptions(0).Top = 840
  picOptions(0).ZOrder 0
  mbDataChanged = False
  
  If CheckNewDB(ADOprimaryrs, "Payroll") = True Then
    ADOprimaryrs.AddNew
  End If
  
GetTextColor Me
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
  
  Me.Width = 9540
  Me.Height = 7725
  
SkipResize:
  frPrimary.Left = (Me.ScaleWidth - frPrimary.Width) / 2
  'Picture1.Left = frprimary.Left
  'lblTop.Width = frPrimary.Width
  frPrimary.Top = (Me.ScaleHeight - frPrimary.Height - picButtons.Height) / 2
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If mbAddNewFlag Then Exit Sub

  Select Case KeyCode
    Case vbKeyEscape
      cmdClose_Click
  End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo FormErr
ShowStatus True
If chkAuto.Value = 1 Then _
    db.Execute "Update [AR SALES] SET [AR SALE Select to Pay]= -1 WHERE [AR SALE Commission Paid] = 0"

    EndLoad db, ADOprimaryrs, "Payroll"
    
ShowStatus False
    If UnloadForm(ADOprimaryrs) = 0 Then
        db.Close
        Set db = Nothing
    Else
        Cancel = 1
    End If
Set frm_SYS_Setup_Payroll = Nothing
Exit Sub
FormErr:
  MsgBox Err.Description
  ShowStatus False
End Sub

Private Sub adoPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This will display the current record position for this recordset
  'lblStatus.Caption = "Record: " & CStr(adoPrimaryRS.AbsolutePosition)
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

Private Sub cmdRefresh_Click()
  'This is only needed for multi user apps
On Error GoTo RefreshErr
    With ADOprimaryrs
        If .EditMode = adEditInProgress Then .CancelUpdate
        .Requery
    End With
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub

Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr
  
  With ADOprimaryrs
        Dim oTxt As TextBox
          For Each oTxt In Me.txtRO
            If oTxt.Text = "" Then
              If ADOprimaryrs("" & oTxt.DataField & "").Type = 203 Or ADOprimaryrs("" & oTxt.DataField & "").Type = 202 Then oTxt.Text = " "
            End If
          Next
  .Update
  .Requery
  End With
  'mbEditFlag = False
  'mbAddNewFlag = False
  'mbDataChanged = False
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub SetButtons(bVal As Boolean)
  cmdAdd.Visible = bVal
  cmdUpdate.Visible = Not bVal
  cmdCancel.Visible = Not bVal
  cmdDelete.Visible = bVal
  cmdClose.Visible = bVal
  cmdRefresh.Visible = bVal
  cmdNext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdLast.Enabled = bVal
  cmdPrevious.Enabled = bVal
End Sub

Private Sub tbPay_Click()
    Dim i As Integer
    'show and enable the selected tab's controls
    'and hide and disable all others
    For i = 0 To tbPay.Tabs.count - 1
        If i = tbPay.SelectedItem.Index - 1 Then
            picOptions(i).Left = 360
            picOptions(i).Top = 840
            picOptions(i).Enabled = True
            picOptions(i).ZOrder 0
        Else
            picOptions(i).Left = -20000
            picOptions(i).Enabled = False
        End If
    Next

End Sub

Private Sub txtRO_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
Case 0, 1, 6, 7, 8, 9, 10, 11, 12, 13
    keyResponse = CtrlValidate(KeyAscii, "0123456789.")
    If keyResponse = True Then
    Else
       KeyAscii = 0
    End If
End Select

End Sub

Private Sub txtRO_LostFocus(Index As Integer)
Select Case Index
Case 0, 8, 11, 12, 13
    txtRO(Index) = Format(txtRO(Index), "00.00")
Case 1, 7, 9, 10
    txtRO(Index) = FormatCurr(txtRO(Index))
Case 6
    txtRO(Index) = Val(txtRO(Index))
End Select

End Sub
