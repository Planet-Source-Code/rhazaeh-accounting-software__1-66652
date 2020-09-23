VERSION 5.00
Begin VB.Form frm_GL_Account_Balances 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Account Balances"
   ClientHeight    =   7365
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   7320
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   7320
   Begin VB.TextBox txtFields 
      DataField       =   "GL COA Account No"
      DataSource      =   "adoPrimaryRS"
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "GL COA Account Name"
      DataSource      =   "adoPrimaryRS"
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   600
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "GL COA CY Beginning Amt"
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
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   1440
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "GL COA CY Period 1 Amt"
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
      Enabled         =   0   'False
      Height          =   285
      Index           =   4
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "GL COA CY Period 10 Amt"
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
      Enabled         =   0   'False
      Height          =   285
      Index           =   5
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   5040
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "GL COA CY Period 11 Amt"
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
      Enabled         =   0   'False
      Height          =   285
      Index           =   6
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   5400
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "GL COA CY Period 12 Amt"
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
      Enabled         =   0   'False
      Height          =   285
      Index           =   7
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   5760
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "GL COA CY Period 13 Amt"
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
      Enabled         =   0   'False
      Height          =   285
      Index           =   8
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   6120
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "GL COA CY Period 2 Amt"
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
      Enabled         =   0   'False
      Height          =   285
      Index           =   10
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "GL COA CY Period 3 Amt"
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
      Enabled         =   0   'False
      Height          =   285
      Index           =   11
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "GL COA CY Period 4 Amt"
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
      Enabled         =   0   'False
      Height          =   285
      Index           =   12
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   2880
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "GL COA CY Period 5 Amt"
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
      Enabled         =   0   'False
      Height          =   285
      Index           =   13
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   3240
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "GL COA CY Period 6 Amt"
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
      Enabled         =   0   'False
      Height          =   285
      Index           =   14
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   3600
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "GL COA CY Period 7 Amt"
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
      Enabled         =   0   'False
      Height          =   285
      Index           =   15
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   3960
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "GL COA CY Period 8 Amt"
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
      Enabled         =   0   'False
      Height          =   285
      Index           =   16
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   4320
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "GL COA CY Period 9 Amt"
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
      Enabled         =   0   'False
      Height          =   285
      Index           =   17
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   4680
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "GL COA PY Beginning Amt"
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
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   1440
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "GL COA PY Period 1 Amt"
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
      Enabled         =   0   'False
      Height          =   285
      Index           =   9
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "GL COA PY Period 10 Amt"
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
      Enabled         =   0   'False
      Height          =   285
      Index           =   18
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   5040
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "GL COA PY Period 11 Amt"
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
      Enabled         =   0   'False
      Height          =   285
      Index           =   19
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   5400
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "GL COA PY Period 12 Amt"
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
      Enabled         =   0   'False
      Height          =   285
      Index           =   20
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   5760
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "GL COA PY Period 13 Amt"
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
      Enabled         =   0   'False
      Height          =   285
      Index           =   21
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   6120
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "GL COA PY Period 2 Amt"
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
      Enabled         =   0   'False
      Height          =   285
      Index           =   22
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "GL COA PY Period 3 Amt"
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
      Enabled         =   0   'False
      Height          =   285
      Index           =   23
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "GL COA PY Period 4 Amt"
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
      Enabled         =   0   'False
      Height          =   285
      Index           =   24
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   2880
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "GL COA PY Period 5 Amt"
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
      Enabled         =   0   'False
      Height          =   285
      Index           =   25
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   3240
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "GL COA PY Period 6 Amt"
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
      Enabled         =   0   'False
      Height          =   285
      Index           =   26
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   3600
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "GL COA PY Period 7 Amt"
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
      Enabled         =   0   'False
      Height          =   285
      Index           =   27
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   3960
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "GL COA PY Period 8 Amt"
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
      Enabled         =   0   'False
      Height          =   285
      Index           =   28
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   4320
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "GL COA PY Period 9 Amt"
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
      Enabled         =   0   'False
      Height          =   285
      Index           =   29
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   4680
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "GL COA BUD Beginning Amt"
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
      Index           =   30
      Left            =   5400
      TabIndex        =   22
      Top             =   1440
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "GL COA BUD Period 1 Amt"
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
      Index           =   31
      Left            =   5400
      TabIndex        =   21
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "GL COA BUD Period 10 Amt"
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
      Index           =   32
      Left            =   5400
      TabIndex        =   20
      Top             =   5040
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "GL COA BUD Period 11 Amt"
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
      Index           =   33
      Left            =   5400
      TabIndex        =   19
      Top             =   5400
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "GL COA BUD Period 12 Amt"
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
      Index           =   34
      Left            =   5400
      TabIndex        =   18
      Top             =   5760
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "GL COA BUD Period 13 Amt"
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
      Index           =   35
      Left            =   5400
      TabIndex        =   17
      Top             =   6120
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "GL COA BUD Period 2 Amt"
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
      Index           =   36
      Left            =   5400
      TabIndex        =   16
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "GL COA BUD Period 3 Amt"
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
      Index           =   37
      Left            =   5400
      TabIndex        =   15
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "GL COA BUD Period 4 Amt"
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
      Index           =   38
      Left            =   5400
      TabIndex        =   14
      Top             =   2880
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "GL COA BUD Period 5 Amt"
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
      Index           =   39
      Left            =   5400
      TabIndex        =   13
      Top             =   3240
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "GL COA BUD Period 6 Amt"
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
      Index           =   40
      Left            =   5400
      TabIndex        =   12
      Top             =   3600
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "GL COA BUD Period 7 Amt"
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
      Index           =   41
      Left            =   5400
      TabIndex        =   11
      Top             =   3960
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "GL COA BUD Period 8 Amt"
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
      Index           =   42
      Left            =   5400
      TabIndex        =   10
      Top             =   4320
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "GL COA BUD Period 9 Amt"
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
      Index           =   43
      Left            =   5400
      TabIndex        =   9
      Top             =   4680
      Width           =   1695
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   7320
      TabIndex        =   6
      Top             =   6765
      Width           =   7320
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   300
         Left            =   1080
         TabIndex        =   8
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
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
      ScaleWidth      =   7320
      TabIndex        =   0
      Top             =   7065
      Width           =   7320
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Picture         =   "frm_GL_Account_Balances.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Picture         =   "frm_GL_Account_Balances.frx":0342
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Picture         =   "frm_GL_Account_Balances.frx":0684
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Picture         =   "frm_GL_Account_Balances.frx":09C6
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
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Account Number"
      Height          =   255
      Index           =   0
      Left            =   -120
      TabIndex        =   71
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Account Name"
      Height          =   255
      Index           =   1
      Left            =   -120
      TabIndex        =   70
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Beginning Balance"
      Height          =   255
      Index           =   3
      Left            =   -120
      TabIndex        =   69
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Period 1"
      Height          =   255
      Index           =   4
      Left            =   -120
      TabIndex        =   68
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Period 10"
      Height          =   255
      Index           =   5
      Left            =   -120
      TabIndex        =   67
      Top             =   5040
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Period 11"
      Height          =   255
      Index           =   6
      Left            =   -120
      TabIndex        =   66
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Period 12"
      Height          =   255
      Index           =   7
      Left            =   -120
      TabIndex        =   65
      Top             =   5760
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Period 13"
      Height          =   255
      Index           =   8
      Left            =   -120
      TabIndex        =   64
      Top             =   6120
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Period 2"
      Height          =   255
      Index           =   10
      Left            =   -120
      TabIndex        =   63
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Period 3"
      Height          =   255
      Index           =   11
      Left            =   -120
      TabIndex        =   62
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Period 4"
      Height          =   255
      Index           =   12
      Left            =   -120
      TabIndex        =   61
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Period 5"
      Height          =   255
      Index           =   13
      Left            =   -120
      TabIndex        =   60
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Period 6"
      Height          =   255
      Index           =   14
      Left            =   -120
      TabIndex        =   59
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Period 7"
      Height          =   255
      Index           =   15
      Left            =   -120
      TabIndex        =   58
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Period 8"
      Height          =   255
      Index           =   16
      Left            =   -120
      TabIndex        =   57
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Period 9"
      Height          =   255
      Index           =   17
      Left            =   -120
      TabIndex        =   56
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      Caption         =   "Current Year"
      DataSource      =   "adoPrimaryRS"
      Height          =   255
      Index           =   2
      Left            =   1800
      TabIndex        =   55
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      Caption         =   "Previous Year"
      DataSource      =   "adoPrimaryRS"
      Height          =   255
      Index           =   9
      Left            =   3600
      TabIndex        =   54
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      Caption         =   "Budget"
      DataSource      =   "adoPrimaryRS"
      Height          =   255
      Index           =   18
      Left            =   5400
      TabIndex        =   53
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   7320
      Y1              =   6720
      Y2              =   6720
   End
End
Attribute VB_Name = "frm_GL_Account_Balances"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ADOprimaryrs As ADODB.Recordset
Attribute ADOprimaryrs.VB_VarHelpID = -1
Dim db As ADODB.Connection
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean

Private Sub Form_Load()
  GetTextColor Me
  mbDataChanged = False
End Sub

Private Sub Form_Resize()
  On Error Resume Next
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

Private Sub cmdRefresh_Click()
  'This is only needed for multi user apps
  On Error GoTo RefreshErr
  
  'adoPrimaryRS.Close
  'Set adoPrimaryRS = New ADODB.Recordset
  
  'If IsNull(gCOAdrill) Then
  '  adoPrimaryRS.Open "select * from [GL Chart Of Accounts] order by [GL COA Account No]", db, adOpenStatic, adLockOptimistic
  'Else
  '  adoPrimaryRS.Open "select * from [GL Chart Of Accounts] where [GL COA Account No] = '" & gCOAdrill & "'", db, adOpenStatic, adLockOptimistic
  'End If

  'Dim oText As TextBox
  'Bind the text boxes to the data provider
  'For Each oText In Me.txtFields
  '  Set oText.DataSource = adoPrimaryRS
  'Next
  
  ADOprimaryrs.Requery
  mbDataChanged = False

  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub

Private Sub cmdCancel_Click()
  ObAddNewFlag = False
  On Error Resume Next

  mbEditFlag = False
  mdoPrimaryRS.CancelUpdate
  If mvBookMark > 0 Then
    ADOprimaryrs.Bookmark = mvBookMark
  Else
    ADOprimaryrs.MoveFirst
  End If
  mbDataChanged = False

End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdFirst_Click()
  On Error GoTo GoFirstError

  ADOprimaryrs.MoveFirst
  mbDataChanged = False

  Exit Sub

GoFirstError:
  MsgBox Err.Description
End Sub

Private Sub cmdLast_Click()
  On Error GoTo GoLastError

  ADOprimaryrs.MoveLast
  mbDataChanged = False

  Exit Sub

GoLastError:
  MsgBox Err.Description
End Sub

Private Sub cmdNext_Click()
  On Error GoTo GoNextError

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
  On Error GoTo GoPrevError

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

Public Sub OpenAccount(AccNum As String)
    Dim sql As String
    Dim SQLW As String
    sql = "select [GL COA Account No],[GL COA Account Name]," & _
        "[GL COA CY Beginning Amt],[GL COA CY Period 1 Amt]," & _
        "[GL COA CY Period 2 Amt],[GL COA CY Period 3 Amt]," & _
        "[GL COA CY Period 4 Amt],[GL COA CY Period 5 Amt]," & _
        "[GL COA CY Period 6 Amt],[GL COA CY Period 7 Amt]," & _
        "[GL COA CY Period 8 Amt],[GL COA CY Period 9 Amt]," & _
        "[GL COA CY Period 10 Amt],[GL COA CY Period 11 Amt]," & _
        "[GL COA CY Period 12 Amt],[GL COA CY Period 13 Amt]," & _
        "[GL COA PY Beginning Amt],[GL COA PY Period 1 Amt]," & _
        "[GL COA PY Period 2 Amt],[GL COA PY Period 3 Amt]," & _
        "[GL COA PY Period 4 Amt],[GL COA PY Period 5 Amt]," & _
        "[GL COA PY Period 6 Amt],[GL COA PY Period 7 Amt]," & _
        "[GL COA PY Period 8 Amt],[GL COA PY Period 9 Amt]," & _
        "[GL COA PY Period 10 Amt],[GL COA PY Period 11 Amt]," & _
        "[GL COA PY Period 12 Amt],[GL COA PY Period 13 Amt]," & _
        "[GL COA BUD Beginning Amt],[GL COA BUD Period 1 Amt]," & _
        "[GL COA BUD Period 2 Amt],[GL COA BUD Period 3 Amt]," & _
        "[GL COA BUD Period 4 Amt],[GL COA BUD Period 5 Amt]," & _
        "[GL COA BUD Period 6 Amt],[GL COA BUD Period 7 Amt]," & _
        "[GL COA BUD Period 8 Amt],[GL COA BUD Period 9 Amt]," & _
        "[GL COA BUD Period 10 Amt],[GL COA BUD Period 11 Amt]," & _
        "[GL COA BUD Period 12 Amt],[GL COA BUD Period 13 Amt] " & _
        "from [GL Chart Of Accounts]"
    SQLW = " where [GL COA Account No] = '" & AccNum & "'"
    
    If db Is Nothing Then
        Set db = New ADODB.Connection
        db.CursorLocation = adUseClient
        db.Open gblADOProvider
    End If

    'closes previously opened recordset
    'If adoPrimaryRS.State Is adStateOpen Then
    '    adoPrimaryRS.Close
    '    Set adoPrimaryRS = Nothing
    'End If
    
    Set ADOprimaryrs = New ADODB.Recordset
    ADOprimaryrs.Open sql & SQLW, db, adOpenStatic, adLockOptimistic
    
    Dim oText As TextBox
    'Bind the text boxes to the data provider
    For Each oText In Me.txtfields
        Set oText.DataSource = ADOprimaryrs
    Next
End Sub
