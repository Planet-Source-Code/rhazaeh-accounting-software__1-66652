VERSION 5.00
Begin VB.Form frm_SYS_Setup_Periods 
   Caption         =   "Setup Accounting Periods"
   ClientHeight    =   6300
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   11460
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6300
   ScaleWidth      =   11460
   Begin VB.Frame frPrimary 
      Height          =   5415
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   11295
      Begin VB.TextBox txtPeriodsClosed 
         DataField       =   " "
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   0
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   54
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtPeriods 
         DataField       =   "SYS COM P1 Date"
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
         Index           =   30
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   53
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtPeriodsClosed 
         DataField       =   " "
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   9
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   52
         Top             =   3720
         Width           =   1335
      End
      Begin VB.TextBox txtPeriods 
         DataField       =   "SYS COM P10 Date"
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
         Index           =   39
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   51
         Top             =   3720
         Width           =   1335
      End
      Begin VB.TextBox txtPeriodsClosed 
         DataField       =   " "
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   10
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   4080
         Width           =   1335
      End
      Begin VB.TextBox txtPeriods 
         DataField       =   "SYS COM P11 Date"
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
         Index           =   40
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   49
         Top             =   4080
         Width           =   1335
      End
      Begin VB.TextBox txtPeriodsClosed 
         DataField       =   " "
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   11
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   48
         Top             =   4440
         Width           =   1335
      End
      Begin VB.TextBox txtPeriods 
         DataField       =   "SYS COM P12 Date"
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
         Index           =   41
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   47
         Top             =   4440
         Width           =   1335
      End
      Begin VB.TextBox txtPeriodsClosed 
         DataField       =   " "
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   12
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   4800
         Width           =   1335
      End
      Begin VB.TextBox txtPeriods 
         DataField       =   "SYS COM P13 Date"
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
         Index           =   42
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   4800
         Width           =   1335
      End
      Begin VB.TextBox txtPeriodsClosed 
         DataField       =   " "
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   1
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtPeriods 
         DataField       =   "SYS COM P2 Date"
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
         Index           =   31
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtPeriodsClosed 
         DataField       =   " "
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   2
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txtPeriods 
         DataField       =   "SYS COM P3 Date"
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
         Index           =   32
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txtPeriodsClosed 
         DataField       =   " "
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   3
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox txtPeriods 
         DataField       =   "SYS COM P4 Date"
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
         Index           =   33
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox txtPeriodsClosed 
         DataField       =   " "
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   4
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox txtPeriods 
         DataField       =   "SYS COM P5 Date"
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
         Index           =   34
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox txtPeriodsClosed 
         DataField       =   " "
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   5
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox txtPeriods 
         DataField       =   "SYS COM P6 Date"
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
         Index           =   35
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox txtPeriodsClosed 
         DataField       =   " "
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   6
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox txtPeriods 
         DataField       =   "SYS COM P7 Date"
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
         Index           =   36
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox txtPeriodsClosed 
         DataField       =   " "
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   7
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   3000
         Width           =   1335
      End
      Begin VB.TextBox txtPeriods 
         DataField       =   "SYS COM P8 Date"
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
         Index           =   37
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   3000
         Width           =   1335
      End
      Begin VB.TextBox txtPeriodsClosed 
         DataField       =   " "
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   8
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   3360
         Width           =   1335
      End
      Begin VB.TextBox txtPeriods 
         DataField       =   "SYS COM P9 Date"
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
         Index           =   38
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Frame frPeriods 
         Caption         =   "AutoSetup"
         Height          =   4695
         Left            =   6600
         TabIndex        =   18
         Top             =   360
         Width           =   4455
         Begin VB.TextBox txtPeriods 
            Alignment       =   2  'Center
            DataField       =   "SYS COM Fiscal Year"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   5
            Left            =   360
            Locked          =   -1  'True
            TabIndex        =   91
            Top             =   1320
            Width           =   1695
         End
         Begin VB.CommandButton cmdDate 
            Height          =   285
            Index           =   0
            Left            =   1680
            Picture         =   "frm_SYS_Setup_Periods.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   90
            Top             =   600
            Width           =   375
         End
         Begin VB.CommandButton cmdDate 
            Height          =   285
            Index           =   1
            Left            =   3720
            Picture         =   "frm_SYS_Setup_Periods.frx":05DA
            Style           =   1  'Graphical
            TabIndex        =   89
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox txtPeriods 
            DataField       =   "SYS COM Fiscal Start Date"
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
            Index           =   26
            Left            =   360
            Locked          =   -1  'True
            TabIndex        =   27
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox txtPeriods 
            DataField       =   "SYS COM Fiscal End Date"
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
            Index           =   27
            Left            =   2400
            Locked          =   -1  'True
            TabIndex        =   26
            Top             =   600
            Width           =   1335
         End
         Begin VB.Frame frPeriodSelector 
            Caption         =   "PeriodSelector"
            Height          =   2175
            Left            =   120
            TabIndex        =   20
            Top             =   2400
            Width           =   4215
            Begin VB.TextBox txtPeriods 
               Alignment       =   2  'Center
               DataField       =   "SYS COM Fiscal Dist Type"
               DataSource      =   "adoPrimaryRS"
               Height          =   285
               Index           =   8
               Left            =   2640
               Locked          =   -1  'True
               TabIndex        =   93
               Top             =   720
               Width           =   1455
            End
            Begin VB.OptionButton optPeriodSelector 
               Caption         =   "I used quaterely Periods"
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   25
               Top             =   480
               Value           =   -1  'True
               Width           =   2055
            End
            Begin VB.OptionButton optPeriodSelector 
               Caption         =   "I use monthly Periods"
               Height          =   255
               Index           =   1
               Left            =   240
               TabIndex        =   24
               Top             =   840
               Width           =   2055
            End
            Begin VB.OptionButton optPeriodSelector 
               Caption         =   "Just divide the year into equal periods"
               Height          =   255
               Index           =   2
               Left            =   240
               TabIndex        =   23
               Top             =   1560
               Width           =   3015
            End
            Begin VB.ComboBox cbPeriodSelector 
               Enabled         =   0   'False
               Height          =   315
               ItemData        =   "frm_SYS_Setup_Periods.frx":0BB4
               Left            =   3240
               List            =   "frm_SYS_Setup_Periods.frx":0BDF
               TabIndex        =   22
               Top             =   1560
               Width           =   855
            End
            Begin VB.OptionButton optPeriodSelector 
               Caption         =   "I rather use yearly Periods"
               Height          =   255
               Index           =   3
               Left            =   240
               TabIndex        =   21
               Top             =   1200
               Width           =   2295
            End
            Begin VB.Label lblLabels 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Fiscal Dist Type"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   3
               Left            =   2640
               TabIndex        =   94
               Top             =   480
               Width           =   1455
            End
         End
         Begin VB.CommandButton cmdAutoSetup 
            Caption         =   "AutoSetup"
            Height          =   1095
            Left            =   3240
            Picture         =   "frm_SYS_Setup_Periods.frx":0C0E
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   1300
            Width           =   1095
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Fiscal Year"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   92
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Last Day Of Fiscal Year"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   2400
            TabIndex        =   88
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "First Day Of Fiscal Year"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   87
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label1 
            Caption         =   "Use the Period selector to choose how you will divide your fiscal year. "
            Height          =   615
            Index           =   1
            Left            =   120
            TabIndex        =   28
            Top             =   1800
            Width           =   3015
         End
      End
      Begin VB.CheckBox ChkPeriods 
         DataField       =   "SYS COM P1 Closed"
         Height          =   255
         Index           =   0
         Left            =   6120
         TabIndex        =   17
         Top             =   480
         Width           =   255
      End
      Begin VB.CheckBox ChkPeriods 
         Height          =   255
         Index           =   1
         Left            =   6120
         TabIndex        =   16
         Top             =   840
         Width           =   255
      End
      Begin VB.CheckBox ChkPeriods 
         Height          =   255
         Index           =   2
         Left            =   6120
         TabIndex        =   15
         Top             =   1200
         Width           =   255
      End
      Begin VB.CheckBox ChkPeriods 
         Height          =   255
         Index           =   3
         Left            =   6120
         TabIndex        =   14
         Top             =   1560
         Width           =   255
      End
      Begin VB.CheckBox ChkPeriods 
         Height          =   255
         Index           =   4
         Left            =   6120
         TabIndex        =   13
         Top             =   1920
         Width           =   255
      End
      Begin VB.CheckBox ChkPeriods 
         Height          =   255
         Index           =   5
         Left            =   6120
         TabIndex        =   12
         Top             =   2280
         Width           =   255
      End
      Begin VB.CheckBox ChkPeriods 
         Height          =   255
         Index           =   6
         Left            =   6120
         TabIndex        =   11
         Top             =   2640
         Width           =   255
      End
      Begin VB.CheckBox ChkPeriods 
         Height          =   255
         Index           =   7
         Left            =   6120
         TabIndex        =   10
         Top             =   3000
         Width           =   255
      End
      Begin VB.CheckBox ChkPeriods 
         Height          =   255
         Index           =   8
         Left            =   6120
         TabIndex        =   9
         Top             =   3360
         Width           =   255
      End
      Begin VB.CheckBox ChkPeriods 
         Height          =   255
         Index           =   9
         Left            =   6120
         TabIndex        =   8
         Top             =   3720
         Width           =   255
      End
      Begin VB.CheckBox ChkPeriods 
         Height          =   255
         Index           =   10
         Left            =   6120
         TabIndex        =   7
         Top             =   4080
         Width           =   255
      End
      Begin VB.CheckBox ChkPeriods 
         Height          =   255
         Index           =   11
         Left            =   6120
         TabIndex        =   6
         Top             =   4440
         Width           =   255
      End
      Begin VB.CheckBox ChkPeriods 
         Height          =   255
         Index           =   12
         Left            =   6120
         TabIndex        =   5
         Top             =   4800
         Width           =   255
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Period 1 Closed"
         Height          =   255
         Index           =   33
         Left            =   3120
         TabIndex        =   80
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Period 1 Date"
         DataSource      =   "adoPrimaryRS"
         Height          =   255
         Index           =   34
         Left            =   240
         TabIndex        =   79
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Period 10 Closed"
         Height          =   255
         Index           =   35
         Left            =   3120
         TabIndex        =   78
         Top             =   3720
         Width           =   1455
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Period 10 Date"
         DataSource      =   "adoPrimaryRS"
         Height          =   255
         Index           =   36
         Left            =   240
         TabIndex        =   77
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Period 11 Closed"
         Height          =   255
         Index           =   37
         Left            =   3120
         TabIndex        =   76
         Top             =   4080
         Width           =   1455
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Period 11 Date"
         DataSource      =   "adoPrimaryRS"
         Height          =   255
         Index           =   38
         Left            =   240
         TabIndex        =   75
         Top             =   4080
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Period 12 Closed"
         Height          =   255
         Index           =   39
         Left            =   3120
         TabIndex        =   74
         Top             =   4440
         Width           =   1455
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Period 12 Date"
         DataSource      =   "adoPrimaryRS"
         Height          =   255
         Index           =   40
         Left            =   240
         TabIndex        =   73
         Top             =   4440
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Period 13 Closed"
         Height          =   255
         Index           =   41
         Left            =   3120
         TabIndex        =   72
         Top             =   4800
         Width           =   1455
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Period 13 Date"
         DataSource      =   "adoPrimaryRS"
         Height          =   255
         Index           =   42
         Left            =   240
         TabIndex        =   71
         Top             =   4800
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Period 2 Closed"
         Height          =   255
         Index           =   43
         Left            =   3120
         TabIndex        =   70
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Period 2 Date"
         DataSource      =   "adoPrimaryRS"
         Height          =   255
         Index           =   44
         Left            =   240
         TabIndex        =   69
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Period 3 Closed"
         Height          =   255
         Index           =   45
         Left            =   3120
         TabIndex        =   68
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Period 3 Date"
         DataSource      =   "adoPrimaryRS"
         Height          =   255
         Index           =   46
         Left            =   240
         TabIndex        =   67
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Period 4 Closed"
         Height          =   255
         Index           =   47
         Left            =   3120
         TabIndex        =   66
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Period 4 Date"
         DataSource      =   "adoPrimaryRS"
         Height          =   255
         Index           =   48
         Left            =   240
         TabIndex        =   65
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Period 5 Closed"
         Height          =   255
         Index           =   49
         Left            =   3120
         TabIndex        =   64
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Period 5 Date"
         DataSource      =   "adoPrimaryRS"
         Height          =   255
         Index           =   50
         Left            =   240
         TabIndex        =   63
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Period 6 Closed"
         Height          =   255
         Index           =   51
         Left            =   3120
         TabIndex        =   62
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Period 6 Date"
         DataSource      =   "adoPrimaryRS"
         Height          =   255
         Index           =   52
         Left            =   240
         TabIndex        =   61
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Period 7 Closed"
         Height          =   255
         Index           =   53
         Left            =   3120
         TabIndex        =   60
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Period 7 Date"
         DataSource      =   "adoPrimaryRS"
         Height          =   255
         Index           =   54
         Left            =   240
         TabIndex        =   59
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Period 8 Closed"
         Height          =   255
         Index           =   55
         Left            =   3120
         TabIndex        =   58
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Period 8 Date"
         DataSource      =   "adoPrimaryRS"
         Height          =   255
         Index           =   56
         Left            =   240
         TabIndex        =   57
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Period 9 Closed"
         Height          =   255
         Index           =   57
         Left            =   3120
         TabIndex        =   56
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Period 9 Date"
         DataSource      =   "adoPrimaryRS"
         Height          =   255
         Index           =   58
         Left            =   240
         TabIndex        =   55
         Top             =   3360
         Width           =   1215
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
      ScaleWidth      =   11460
      TabIndex        =   0
      Top             =   6000
      Width           =   11460
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   300
         Left            =   2280
         TabIndex        =   2
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   300
         Left            =   1200
         TabIndex        =   1
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   300
         Left            =   120
         TabIndex        =   3
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.PictureBox picStatBox 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   11460
      TabIndex        =   81
      Top             =   6870
      Visible         =   0   'False
      Width           =   11460
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Picture         =   "frm_SYS_Setup_Periods.frx":0F18
         Style           =   1  'Graphical
         TabIndex        =   85
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Picture         =   "frm_SYS_Setup_Periods.frx":125A
         Style           =   1  'Graphical
         TabIndex        =   84
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   10680
         Picture         =   "frm_SYS_Setup_Periods.frx":159C
         Style           =   1  'Graphical
         TabIndex        =   83
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   11025
         Picture         =   "frm_SYS_Setup_Periods.frx":18DE
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   86
         Top             =   0
         Width           =   9960
      End
   End
   Begin VB.Label lblTop 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Accounting Periods"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   20
      TabIndex        =   95
      Top             =   120
      Width           =   11415
   End
End
Attribute VB_Name = "frm_SYS_Setup_Periods"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ADOprimaryrs As ADODB.Recordset
Attribute ADOprimaryrs.VB_VarHelpID = -1
Dim db As ADODB.Connection
Dim NewLoad As Boolean

Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean

Private Function GetDate() As Boolean
Dim i As Integer

  'Automatically fill in dates based on Fiscal Start Date and Fiscal End Date
  If Not IsDate(txtPeriods(26)) Then
    MsgBox "Please enter first day of fiscal year!", , "Error"
    GetDate = False
    Exit Function
  Else
    txtPeriods(26) = FormatDate(txtPeriods(26))
  End If
  
  If Not IsDate(txtPeriods(27)) Then
    MsgBox "Please enter last day of fiscal year!", , "Error"
    GetDate = False
    Exit Function
  Else
    txtPeriods(27) = FormatDate(txtPeriods(27))
  End If
  
  For i = 30 To 42
    txtPeriods(i) = ""
  Next
  
  For i = 0 To 12
    txtPeriodsClosed(i) = ""
  Next
  
  'DateADD--Returns a Variant (Date) containing a date to which a specified time interval has been added. Search MSDN
    
  'if quarterly
    If optPeriodSelector(0).Value = True Then
        txtPeriods(30) = txtPeriods(26)
        For i = 31 To 33
            txtPeriods(i) = FormatDate(DateAdd("q", 1, txtPeriods(i - 1)))
            txtPeriodsClosed(i - 31) = FormatDate(DateAdd("q", 1, txtPeriods(i - 1)) - 1)
        Next
        txtPeriodsClosed(3) = FormatDate(DateAdd("q", 1, txtPeriods(i - 1)) - 1)
    'if monthly
    ElseIf optPeriodSelector(1).Value = True Then
        txtPeriods(30) = txtPeriods(26)
        For i = 31 To 41
            txtPeriods(i) = FormatDate(DateAdd("m", 1, txtPeriods(i - 1)))
            txtPeriodsClosed(i - 31) = Format(DateAdd("m", 1, txtPeriods(i - 1)) - 1)
        Next
        txtPeriodsClosed(i - 31) = FormatDate(DateAdd("m", 1, txtPeriods(i - 1)) - 1)
    'custom
    ElseIf optPeriodSelector(2).Value = True Then
        txtPeriods(30) = FormatDate(txtPeriods(26))
        Dim dtemp As Variant
        Dim strtemp As String
        Dim intCount As Integer
        Dim MinsApart As Double
        Dim MinsSpread As Double
        Dim InterValPeriods As String
        
        InterValPeriods = "d"
        MinsApart = DateDiff(InterValPeriods, txtPeriods(26), txtPeriods(27))
        cbPeriodSelector.Text = Val(cbPeriodSelector.Text)
        If cbPeriodSelector.Text = "" Or cbPeriodSelector.Text <= 0 Or cbPeriodSelector.Text > 13 Then
            MsgBox "You forgot to select a period but now it's set to 12", vbInformation, "Error Selection"
            cbPeriodSelector.Text = "12"
        End If
        MinsSpread = (MinsApart / Int(cbPeriodSelector.Text))
        dtemp = txtPeriods(30)
        For intCount = 0 To Int(cbPeriodSelector.Text) - 2
            'strtemp = "SYS COM P" & Trim(Str(intCount)) & " Date"
            txtPeriods(31 + intCount) = FormatDate(DateAdd(InterValPeriods, MinsSpread, dtemp))
            txtPeriodsClosed(intCount) = FormatDate(DateAdd(InterValPeriods, MinsSpread, dtemp) - 1)
            dtemp = txtPeriods(31 + intCount)
            'Me(strtemp) = dtemp
        Next
            txtPeriodsClosed(intCount) = FormatDate(txtPeriods(27))
    'yearly
    Else
        txtPeriods(30) = FormatDate(txtPeriods(26))
        txtPeriodsClosed(0) = FormatDate(txtPeriods(27))
    End If
    ADOprimaryrs![SYS COM P1 Date] = txtPeriods(30)
    GetDate = True
End Function


Private Sub cbPeriodSelector_Click()
    GetDate
End Sub

Private Sub cmdAutoSetup_Click()
    GetDate
End Sub

Private Sub cmdDate_Click(Index As Integer)
Select Case Index
Case 0
    Menu_Calendar.WhoCallMe True, 1003
Case 1
    Menu_Calendar.WhoCallMe True, 1004
End Select
'Menu_Calendar.Show vbModal

txtPeriods(5) = Format(txtPeriods(26), "yyyy")
End Sub

Private Sub Form_Load()
ShowStatus True
On Error GoTo FormErr
NewLoad = True
  Set db = New ADODB.Connection
  db.CursorLocation = adUseClient
  db.Open gblADOProvider

  Set ADOprimaryrs = New ADODB.Recordset
  ADOprimaryrs.Open "select [SYS COM Fiscal Start Date],[SYS COM Fiscal End Date],[SYS COM P1 Closed],[SYS COM P1 Date],[SYS COM P10 Closed],[SYS COM P10 Date],[SYS COM P11 Closed],[SYS COM P11 Date],[SYS COM P12 Closed],[SYS COM P12 Date],[SYS COM P13 Closed],[SYS COM P13 Date],[SYS COM P2 Closed],[SYS COM P2 Date],[SYS COM P3 Closed],[SYS COM P3 Date],[SYS COM P4 Closed],[SYS COM P4 Date],[SYS COM P5 Closed],[SYS COM P5 Date],[SYS COM P6 Closed],[SYS COM P6 Date],[SYS COM P7 Closed],[SYS COM P7 Date],[SYS COM P8 Closed],[SYS COM P8 Date],[SYS COM P9 Closed],[SYS COM P9 Date],[SYS COM Fiscal Dist Type],[SYS COM Fiscal Year] from [SYS Company]", db, adOpenStatic, adLockOptimistic

  Dim chk As CheckBox
  Dim txt As TextBox
  'Bind the text boxes to the data provider
  For Each chk In Me.ChkPeriods
        Set chk.DataSource = ADOprimaryrs
  Next
  For Each txt In Me.txtPeriods
        Set txt.DataSource = ADOprimaryrs
  Next
  
  If CheckNewDB(ADOprimaryrs, "Periods") = True Then
    ADOprimaryrs.AddNew
    optPeriodSelector(0).Value = True
  Else
    Select Case ADOprimaryrs![SYS COM Fiscal Dist Type]
      Case "Quartely"
          optPeriodSelector(0).Value = True
      Case "Monthly"
          optPeriodSelector(1).Value = True
      Case "Custom"
          optPeriodSelector(2).Value = True
      Case "Yearly"
          optPeriodSelector(3).Value = True
    End Select
  End If
  NewLoad = False
  GetDate
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
  
  Me.Width = 11550
  Me.Height = 6675
  
SkipResize:
  frPrimary.Left = (Me.ScaleWidth - frPrimary.Width) / 2
  lblTop.Left = frPrimary.Left
  lblTop.Width = frPrimary.Width
  frPrimary.Top = (Me.ScaleHeight - frPrimary.Height - picButtons.Height) / 2 + 230
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
  Set frm_SYS_Setup_Periods = Nothing
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
  End With

  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
  'On Error GoTo DeleteErr
  With ADOprimaryrs
    .Delete
    .MoveNext
    If .EOF Then .MoveLast
  End With
  Exit Sub
DeleteErr:
  MsgBox Err.Description
End Sub

Private Sub cmdRefresh_Click()
  'This is only needed for multi user apps
  'On Error GoTo RefreshErr
  If GetDate = False Then Exit Sub
    With ADOprimaryrs
        If .EditMode <> 0 Then .CancelUpdate
        If .EditMode = 0 Then .Requery
    End With
  Exit Sub
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub

Private Sub cmdEdit_Click()
  'On Error GoTo EditErr

  lblStatus.Caption = "Edit record"
  mbEditFlag = True
  Exit Sub

EditErr:
  MsgBox Err.Description
End Sub
Private Sub cmdCancel_Click()
  'On Error Resume Next

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
  'On Error GoTo UpdateErr
  If GetDate = False Then Exit Sub
  With ADOprimaryrs
    .Update
'  .Requery
  End With
  Exit Sub

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

Private Sub optPeriodSelector_Click(Index As Integer)
cbPeriodSelector.Text = cbPeriodSelector.List(0)
Select Case Index
Case 0
    txtPeriods(8) = "Quartely"
    cbPeriodSelector.Enabled = False
Case 1
    txtPeriods(8) = "Monthly"
    cbPeriodSelector.Enabled = False
Case 2
    txtPeriods(8) = "Custom"
    cbPeriodSelector.Enabled = True
Case 3
    txtPeriods(8) = "Yearly"
    cbPeriodSelector.Enabled = False
End Select
    If NewLoad = False Then txtPeriods(8).SetFocus
    GetDate
End Sub
