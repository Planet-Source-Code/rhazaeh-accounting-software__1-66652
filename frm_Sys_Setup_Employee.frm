VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Begin VB.Form frm_SYS_Setup_Employee 
   Caption         =   "Employee Setup"
   ClientHeight    =   8550
   ClientLeft      =   2220
   ClientTop       =   1950
   ClientWidth     =   11220
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8550
   ScaleWidth      =   11220
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   11220
      TabIndex        =   177
      Top             =   8250
      Width           =   11220
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   300
         Index           =   0
         Left            =   5160
         TabIndex        =   178
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Enabled         =   0   'False
         Height          =   300
         Left            =   4080
         TabIndex        =   179
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Enabled         =   0   'False
         Height          =   300
         Left            =   3000
         TabIndex        =   180
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdCreatePyrll 
         Caption         =   "Create Payroll"
         Height          =   300
         Left            =   1680
         TabIndex        =   181
         Top             =   0
         Width           =   1335
      End
      Begin VB.CommandButton cmdBeginning 
         Caption         =   "Beginning Balances"
         Height          =   300
         Left            =   0
         TabIndex        =   182
         Top             =   0
         Width           =   1695
      End
   End
   Begin VB.Frame frPrimary 
      Height          =   7695
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   11175
      Begin VB.Frame frEmployeeSetup 
         Height          =   7215
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   10935
         Begin VB.Frame Frame1 
            Height          =   1095
            Left            =   50
            TabIndex        =   5
            Top             =   120
            Width           =   3855
            Begin VB.CommandButton cmdEmpID 
               Height          =   285
               Left            =   3240
               Picture         =   "frm_Sys_Setup_Employee.frx":0000
               Style           =   1  'Graphical
               TabIndex        =   9
               Top             =   720
               Width           =   375
            End
            Begin VB.TextBox txtFields 
               DataField       =   "EMP ID"
               DataSource      =   "adoPrimaryRS"
               Height          =   285
               Index           =   1
               Left            =   1200
               TabIndex        =   8
               Top             =   720
               Width           =   2055
            End
            Begin VB.CommandButton cmdAdd 
               Caption         =   "Add"
               Height          =   375
               Left            =   1320
               TabIndex        =   7
               Top             =   240
               Width           =   1095
            End
            Begin VB.CommandButton cmdDelete 
               Caption         =   "Delete"
               Height          =   375
               Left            =   2520
               TabIndex        =   6
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label lblLabels 
               Alignment       =   1  'Right Justify
               Caption         =   "Employee ID:"
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   11
               Top             =   720
               Width           =   975
            End
            Begin VB.Label lblEmp 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   10
               Top             =   1200
               Visible         =   0   'False
               Width           =   3495
            End
         End
         Begin MSDataGridLib.DataGrid grdDataGrid 
            Bindings        =   "frm_Sys_Setup_Employee.frx":014A
            Height          =   1575
            Left            =   120
            TabIndex        =   12
            Top             =   5520
            Width           =   10695
            _ExtentX        =   18865
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
            ColumnCount     =   9
            BeginProperty Column00 
               DataField       =   "ApplyItem"
               Caption         =   "Apply"
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
               DataField       =   "Description"
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
            BeginProperty Column02 
               DataField       =   "ItemAmount"
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
               DataField       =   "ItemPercent"
               Caption         =   "%"
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
               DataField       =   "Basis"
               Caption         =   "Basis"
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
               DataField       =   "WageLow"
               Caption         =   "WageLow"
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
               DataField       =   "WageHigh"
               Caption         =   "WageHigh"
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
               DataField       =   "YTDMax"
               Caption         =   "YTDMax"
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
            BeginProperty Column08 
               DataField       =   "Account"
               Caption         =   "Account"
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
                  Button          =   -1  'True
                  ColumnWidth     =   764.787
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   2250.142
               EndProperty
               BeginProperty Column02 
                  Alignment       =   1
                  ColumnWidth     =   1140.095
               EndProperty
               BeginProperty Column03 
                  Alignment       =   2
                  ColumnWidth     =   675.213
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   959.811
               EndProperty
               BeginProperty Column05 
                  Alignment       =   1
                  ColumnWidth     =   1049.953
               EndProperty
               BeginProperty Column06 
                  Alignment       =   1
                  ColumnWidth     =   1049.953
               EndProperty
               BeginProperty Column07 
                  Alignment       =   1
                  ColumnWidth     =   1049.953
               EndProperty
               BeginProperty Column08 
                  ColumnWidth     =   1244.976
               EndProperty
            EndProperty
         End
         Begin VB.Frame Frame2 
            Height          =   4215
            Left            =   50
            TabIndex        =   38
            Top             =   1200
            Width           =   3855
            Begin VB.TextBox txtFields 
               DataField       =   "EMP Name"
               DataSource      =   "adoPrimaryRS"
               Height          =   285
               Index           =   0
               Left            =   1200
               TabIndex        =   48
               Top             =   480
               Width           =   2535
            End
            Begin VB.TextBox txtFields 
               DataField       =   "EMP Address 1"
               DataSource      =   "adoPrimaryRS"
               Height          =   285
               Index           =   2
               Left            =   1200
               TabIndex        =   47
               Top             =   1200
               Width           =   2535
            End
            Begin VB.TextBox txtFields 
               DataField       =   "EMP Address 2"
               DataSource      =   "adoPrimaryRS"
               Height          =   285
               Index           =   3
               Left            =   1200
               TabIndex        =   46
               Top             =   1560
               Width           =   2535
            End
            Begin VB.TextBox txtFields 
               DataField       =   "EMP City"
               DataSource      =   "adoPrimaryRS"
               Height          =   285
               Index           =   4
               Left            =   1200
               TabIndex        =   45
               Top             =   1920
               Width           =   1575
            End
            Begin VB.TextBox txtFields 
               DataField       =   "EMP State"
               DataSource      =   "adoPrimaryRS"
               Height          =   285
               Index           =   5
               Left            =   1200
               TabIndex        =   44
               Top             =   2280
               Width           =   855
            End
            Begin VB.TextBox txtFields 
               DataField       =   "EMP Postal"
               DataSource      =   "adoPrimaryRS"
               Height          =   285
               Index           =   6
               Left            =   2760
               TabIndex        =   43
               Top             =   2280
               Width           =   975
            End
            Begin VB.TextBox txtFields 
               DataField       =   "EMP Country"
               DataSource      =   "adoPrimaryRS"
               Height          =   285
               Index           =   7
               Left            =   1200
               TabIndex        =   42
               Top             =   2640
               Width           =   1575
            End
            Begin VB.TextBox txtFields 
               DataField       =   "EMP Phone"
               DataSource      =   "adoPrimaryRS"
               Height          =   285
               Index           =   8
               Left            =   1200
               TabIndex        =   41
               Top             =   3000
               Width           =   1575
            End
            Begin VB.TextBox txtFields 
               DataField       =   "EMP Home Fax"
               DataSource      =   "adoPrimaryRS"
               Height          =   285
               Index           =   9
               Left            =   1200
               TabIndex        =   40
               Top             =   3360
               Width           =   1575
            End
            Begin VB.ComboBox Combo1 
               DataField       =   "MALEFEMALE"
               Height          =   315
               ItemData        =   "frm_Sys_Setup_Employee.frx":015A
               Left            =   1200
               List            =   "frm_Sys_Setup_Employee.frx":0164
               Style           =   2  'Dropdown List
               TabIndex        =   39
               Top             =   840
               Width           =   1455
            End
            Begin VB.Label lblLabels 
               Alignment       =   1  'Right Justify
               Caption         =   "Name:"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   57
               Top             =   480
               Width           =   975
            End
            Begin VB.Label lblLabels 
               Alignment       =   1  'Right Justify
               Caption         =   "Address:"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   56
               Top             =   1200
               Width           =   975
            End
            Begin VB.Label lblLabels 
               Alignment       =   1  'Right Justify
               Caption         =   "City:"
               Height          =   255
               Index           =   4
               Left            =   120
               TabIndex        =   55
               Top             =   1920
               Width           =   975
            End
            Begin VB.Label lblLabels 
               Alignment       =   1  'Right Justify
               Caption         =   "State:"
               Height          =   255
               Index           =   5
               Left            =   120
               TabIndex        =   54
               Top             =   1920
               Width           =   975
            End
            Begin VB.Label lblLabels 
               Alignment       =   1  'Right Justify
               Caption         =   "Postal:"
               Height          =   255
               Index           =   6
               Left            =   2160
               TabIndex        =   53
               Top             =   2280
               Width           =   495
            End
            Begin VB.Label lblLabels 
               Alignment       =   1  'Right Justify
               Caption         =   "Country:"
               Height          =   255
               Index           =   7
               Left            =   120
               TabIndex        =   52
               Top             =   2640
               Width           =   975
            End
            Begin VB.Label lblLabels 
               Alignment       =   1  'Right Justify
               Caption         =   "Phone:"
               Height          =   255
               Index           =   8
               Left            =   120
               TabIndex        =   51
               Top             =   3000
               Width           =   975
            End
            Begin VB.Label lblLabels 
               Alignment       =   1  'Right Justify
               Caption         =   "Home Fax:"
               Height          =   255
               Index           =   9
               Left            =   120
               TabIndex        =   50
               Top             =   3360
               Width           =   975
            End
            Begin VB.Label lblLabels 
               Alignment       =   1  'Right Justify
               Caption         =   "Gender/Sex:"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   49
               Top             =   840
               Width           =   975
            End
         End
         Begin VB.PictureBox picOptions 
            BorderStyle     =   0  'None
            Height          =   4740
            Index           =   2
            Left            =   4080
            ScaleHeight     =   4816.452
            ScaleMode       =   0  'User
            ScaleWidth      =   6837.16
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   600
            Width           =   6765
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Status"
               DataField       =   "EMP Inactive YN"
               Height          =   255
               Left            =   4920
               TabIndex        =   192
               Top             =   600
               Width           =   1575
            End
            Begin VB.TextBox txtFields 
               DataField       =   "EMP Notes"
               Height          =   1695
               Index           =   14
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   191
               Top             =   2880
               Width           =   6375
            End
            Begin VB.TextBox txtFields 
               DataField       =   "EMP Emergency Number"
               BeginProperty DataFormat 
                  Type            =   0
                  Format          =   "MM/dd/yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
               DataSource      =   "adoPrimaryRS"
               Height          =   285
               Index           =   13
               Left            =   2520
               TabIndex        =   190
               Top             =   2040
               Width           =   1575
            End
            Begin VB.TextBox txtFields 
               DataField       =   "EMP Emergency Contact"
               BeginProperty DataFormat 
                  Type            =   0
                  Format          =   "MM/dd/yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
               DataSource      =   "adoPrimaryRS"
               Height          =   285
               Index           =   12
               Left            =   2520
               TabIndex        =   189
               Top             =   1560
               Width           =   1575
            End
            Begin VB.CommandButton cmdDate 
               Height          =   285
               Index           =   3
               Left            =   3720
               Picture         =   "frm_Sys_Setup_Employee.frx":0176
               Style           =   1  'Graphical
               TabIndex        =   186
               Top             =   600
               Width           =   375
            End
            Begin VB.CommandButton cmdDate 
               Height          =   285
               Index           =   0
               Left            =   3720
               Picture         =   "frm_Sys_Setup_Employee.frx":0750
               Style           =   1  'Graphical
               TabIndex        =   184
               Top             =   1080
               Width           =   375
            End
            Begin VB.TextBox txtFields 
               DataField       =   "EMP Hire Date"
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
               Index           =   10
               Left            =   2520
               Locked          =   -1  'True
               TabIndex        =   183
               Top             =   1080
               Width           =   1215
            End
            Begin VB.TextBox txtFields 
               DataField       =   "EMP Birthday"
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
               Index           =   11
               Left            =   2520
               TabIndex        =   3
               Top             =   600
               Width           =   1215
            End
            Begin VB.Label lblAE 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Notes"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   10
               Left            =   120
               TabIndex        =   193
               Top             =   2640
               Width           =   6375
            End
            Begin VB.Label lblAE 
               Alignment       =   1  'Right Justify
               Caption         =   "Emergency Contact Number:"
               Height          =   255
               Index           =   9
               Left            =   240
               TabIndex        =   188
               Top             =   2040
               Width           =   2175
            End
            Begin VB.Label lblAE 
               Alignment       =   1  'Right Justify
               Caption         =   "Emergency Contact Person:"
               Height          =   255
               Index           =   8
               Left            =   240
               TabIndex        =   187
               Top             =   1560
               Width           =   2175
            End
            Begin VB.Label lblLabels 
               Alignment       =   1  'Right Justify
               Caption         =   "Hire Date:"
               Height          =   255
               Index           =   10
               Left            =   1440
               TabIndex        =   185
               Top             =   1080
               Width           =   975
            End
            Begin VB.Label lblAE 
               Alignment       =   1  'Right Justify
               Caption         =   "Date of Birth:"
               Height          =   255
               Index           =   11
               Left            =   1200
               TabIndex        =   4
               Top             =   600
               Width           =   1215
            End
         End
         Begin VB.PictureBox picOptions 
            BorderStyle     =   0  'None
            Height          =   4740
            Index           =   0
            Left            =   4080
            ScaleHeight     =   4816.452
            ScaleMode       =   0  'User
            ScaleWidth      =   6837.16
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   600
            Width           =   6765
            Begin VB.ComboBox cbAE 
               DataField       =   "DEPARTMENT"
               Height          =   315
               Index           =   0
               ItemData        =   "frm_Sys_Setup_Employee.frx":0D2A
               Left            =   1320
               List            =   "frm_Sys_Setup_Employee.frx":0D3A
               TabIndex        =   70
               Top             =   600
               Width           =   2655
            End
            Begin VB.ComboBox cbAE 
               DataField       =   "LOCATION"
               Height          =   315
               Index           =   1
               ItemData        =   "frm_Sys_Setup_Employee.frx":0D64
               Left            =   1320
               List            =   "frm_Sys_Setup_Employee.frx":0D6E
               TabIndex        =   69
               Top             =   960
               Width           =   2655
            End
            Begin VB.ComboBox cbAE 
               DataField       =   "PAYTYPE"
               Height          =   315
               Index           =   2
               ItemData        =   "frm_Sys_Setup_Employee.frx":0D8A
               Left            =   1320
               List            =   "frm_Sys_Setup_Employee.frx":0D9A
               TabIndex        =   68
               Top             =   1320
               Width           =   2655
            End
            Begin VB.ComboBox cbAE 
               DataField       =   "PAYFREQUENCY"
               Height          =   315
               Index           =   3
               ItemData        =   "frm_Sys_Setup_Employee.frx":0DCF
               Left            =   1320
               List            =   "frm_Sys_Setup_Employee.frx":0DDF
               TabIndex        =   67
               Top             =   1680
               Width           =   2655
            End
            Begin VB.TextBox txtAE 
               DataField       =   "SALARY"
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
               Index           =   0
               Left            =   1320
               TabIndex        =   66
               Top             =   2280
               Visible         =   0   'False
               Width           =   1695
            End
            Begin VB.TextBox txtAE 
               DataField       =   "HOURLYRATE"
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
               Index           =   2
               Left            =   1320
               TabIndex        =   65
               Top             =   3240
               Visible         =   0   'False
               Width           =   1695
            End
            Begin VB.TextBox txtAE 
               DataField       =   "OTRATE"
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
               Index           =   3
               Left            =   1320
               TabIndex        =   64
               Top             =   3720
               Visible         =   0   'False
               Width           =   1695
            End
            Begin VB.PictureBox picCommision 
               BorderStyle     =   0  'None
               Height          =   495
               Left            =   0
               ScaleHeight     =   495
               ScaleWidth      =   6495
               TabIndex        =   59
               Top             =   2640
               Visible         =   0   'False
               Width           =   6495
               Begin VB.OptionButton optAE 
                  Caption         =   "Sales"
                  Height          =   255
                  Index           =   1
                  Left            =   4800
                  TabIndex        =   62
                  Top             =   120
                  Width           =   1215
               End
               Begin VB.OptionButton optAE 
                  Caption         =   "Profit"
                  Height          =   255
                  Index           =   0
                  Left            =   3360
                  TabIndex        =   61
                  Top             =   120
                  Width           =   1215
               End
               Begin VB.TextBox txtAE 
                  DataField       =   "COMMPERCENT"
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
                  Left            =   1320
                  TabIndex        =   60
                  Top             =   120
                  Width           =   1695
               End
               Begin VB.Label lblAE 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Commision (%):"
                  Height          =   255
                  Index           =   5
                  Left            =   0
                  TabIndex        =   63
                  Top             =   120
                  Width           =   1215
               End
            End
            Begin VB.Label lblAE 
               Alignment       =   1  'Right Justify
               Caption         =   "Department:"
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   77
               Top             =   600
               Width           =   1215
            End
            Begin VB.Label lblAE 
               Alignment       =   1  'Right Justify
               Caption         =   "Location:"
               Height          =   255
               Index           =   1
               Left            =   0
               TabIndex        =   76
               Top             =   960
               Width           =   1215
            End
            Begin VB.Label lblAE 
               Alignment       =   1  'Right Justify
               Caption         =   "Pay Type:"
               Height          =   255
               Index           =   2
               Left            =   0
               TabIndex        =   75
               Top             =   1320
               Width           =   1215
            End
            Begin VB.Label lblAE 
               Alignment       =   1  'Right Justify
               Caption         =   "Pay Frequency:"
               Height          =   255
               Index           =   3
               Left            =   0
               TabIndex        =   74
               Top             =   1680
               Width           =   1215
            End
            Begin VB.Label lblAE 
               Alignment       =   1  'Right Justify
               Caption         =   "Annual Salary:"
               Height          =   255
               Index           =   4
               Left            =   0
               TabIndex        =   73
               Top             =   2280
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.Label lblAE 
               Alignment       =   1  'Right Justify
               Caption         =   "Hourly Rate:"
               Height          =   255
               Index           =   6
               Left            =   0
               TabIndex        =   72
               Top             =   3240
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.Label lblAE 
               Alignment       =   1  'Right Justify
               Caption         =   "Overtime Rate:"
               Height          =   255
               Index           =   7
               Left            =   0
               TabIndex        =   71
               Top             =   3720
               Visible         =   0   'False
               Width           =   1215
            End
         End
         Begin VB.PictureBox picOptions 
            BorderStyle     =   0  'None
            Height          =   4740
            Index           =   1
            Left            =   4080
            ScaleHeight     =   4816.452
            ScaleMode       =   0  'User
            ScaleWidth      =   6837.16
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   600
            Width           =   6765
            Begin VB.TextBox txtDT 
               DataField       =   "YTDLOCAL"
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
               Left            =   4800
               Locked          =   -1  'True
               TabIndex        =   29
               Top             =   3600
               Width           =   1335
            End
            Begin VB.TextBox txtDT 
               DataField       =   "YTDSTATETAX"
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
               Index           =   6
               Left            =   4800
               Locked          =   -1  'True
               TabIndex        =   28
               Top             =   3120
               Width           =   1335
            End
            Begin VB.TextBox txtDT 
               DataField       =   "YTDFIT"
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
               Index           =   5
               Left            =   1680
               Locked          =   -1  'True
               TabIndex        =   27
               Top             =   4080
               Width           =   1335
            End
            Begin VB.TextBox txtDT 
               DataField       =   "YTDFICA"
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
               Index           =   4
               Left            =   1680
               Locked          =   -1  'True
               TabIndex        =   26
               Top             =   3600
               Width           =   1335
            End
            Begin VB.TextBox txtDT 
               DataField       =   "YTDGROSS"
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
               Index           =   3
               Left            =   1680
               Locked          =   -1  'True
               TabIndex        =   25
               Top             =   3120
               Width           =   1335
            End
            Begin VB.CheckBox chkAE 
               Caption         =   "FICA"
               DataField       =   "FICAYN"
               Height          =   255
               Index           =   1
               Left            =   3960
               TabIndex        =   24
               Top             =   2280
               Width           =   1575
            End
            Begin VB.CheckBox chkAE 
               Caption         =   "FIT"
               DataField       =   "FITYN"
               Height          =   255
               Index           =   0
               Left            =   2280
               TabIndex        =   23
               Top             =   2280
               Width           =   1575
            End
            Begin VB.Frame frTab 
               Height          =   1665
               Index           =   1
               Left            =   0
               TabIndex        =   14
               Tag             =   "1061"
               Top             =   480
               Width           =   6720
               Begin VB.TextBox txtDT 
                  DataField       =   "FEDWITHHOLDAMT"
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
                  Index           =   2
                  Left            =   5400
                  TabIndex        =   18
                  Top             =   1200
                  Width           =   1095
               End
               Begin VB.ComboBox cbDT 
                  DataField       =   "FEDFILINGSTATUS"
                  Height          =   315
                  ItemData        =   "frm_Sys_Setup_Employee.frx":0E0B
                  Left            =   1200
                  List            =   "frm_Sys_Setup_Employee.frx":0E15
                  TabIndex        =   17
                  Top             =   720
                  Width           =   2175
               End
               Begin VB.TextBox txtDT 
                  DataField       =   "FEDALLOW"
                  Height          =   285
                  Index           =   1
                  Left            =   5880
                  TabIndex        =   16
                  Top             =   360
                  Width           =   615
               End
               Begin VB.TextBox txtDT 
                  DataField       =   "SS#"
                  Height          =   285
                  Index           =   0
                  Left            =   1200
                  TabIndex        =   15
                  Top             =   360
                  Width           =   2175
               End
               Begin VB.Label lblDT 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Filing Status:"
                  Height          =   255
                  Index           =   10
                  Left            =   120
                  TabIndex        =   22
                  Top             =   720
                  Width           =   975
               End
               Begin VB.Label lblDT 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Additional Amount to Withhold per Pay Period:"
                  Height          =   255
                  Index           =   2
                  Left            =   360
                  TabIndex        =   21
                  Top             =   1200
                  Width           =   4935
               End
               Begin VB.Label lblDT 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Withholding Allowances:"
                  Height          =   255
                  Index           =   1
                  Left            =   3720
                  TabIndex        =   20
                  Top             =   360
                  Width           =   2055
               End
               Begin VB.Label lblDT 
                  Alignment       =   1  'Right Justify
                  Caption         =   "SS#:"
                  Height          =   255
                  Index           =   0
                  Left            =   480
                  TabIndex        =   19
                  Top             =   360
                  Width           =   615
               End
            End
            Begin VB.Label lblDT 
               Alignment       =   1  'Right Justify
               Caption         =   "YTD Local:"
               Height          =   255
               Index           =   9
               Left            =   3600
               TabIndex        =   36
               Top             =   3600
               Width           =   1095
            End
            Begin VB.Label lblDT 
               Alignment       =   1  'Right Justify
               Caption         =   "YTD State:"
               Height          =   255
               Index           =   8
               Left            =   3600
               TabIndex        =   35
               Top             =   3120
               Width           =   1095
            End
            Begin VB.Label lblDT 
               Alignment       =   1  'Right Justify
               Caption         =   "YTD FIT:"
               Height          =   255
               Index           =   7
               Left            =   480
               TabIndex        =   34
               Top             =   4080
               Width           =   1095
            End
            Begin VB.Label lblDT 
               Alignment       =   1  'Right Justify
               Caption         =   "YTD FICA:"
               Height          =   255
               Index           =   6
               Left            =   480
               TabIndex        =   33
               Top             =   3600
               Width           =   1095
            End
            Begin VB.Label lblDT 
               Alignment       =   1  'Right Justify
               Caption         =   "YTD Gross:"
               Height          =   255
               Index           =   5
               Left            =   480
               TabIndex        =   32
               Top             =   3120
               Width           =   1095
            End
            Begin VB.Label lblDT 
               Alignment       =   1  'Right Justify
               Caption         =   "Apply:"
               Height          =   255
               Index           =   3
               Left            =   0
               TabIndex        =   31
               Top             =   2280
               Width           =   2055
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000001&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Balances"
               ForeColor       =   &H0000FFFF&
               Height          =   255
               Index           =   1
               Left            =   0
               TabIndex        =   30
               Top             =   2640
               Width           =   6735
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "W-2"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   0
               Left            =   120
               TabIndex        =   37
               Top             =   120
               Width           =   570
            End
         End
         Begin ComctlLib.TabStrip tbEmployee 
            Height          =   5175
            Left            =   3915
            TabIndex        =   78
            Top             =   240
            Width           =   6975
            _ExtentX        =   12303
            _ExtentY        =   9128
            _Version        =   327682
            BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
               NumTabs         =   3
               BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
                  Caption         =   "Additions/Earnings"
                  Key             =   "ae"
                  Object.Tag             =   ""
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
                  Caption         =   "Deductions/Taxes"
                  Key             =   "dt"
                  Object.Tag             =   ""
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
                  Caption         =   "Emp. Personal Data"
                  Key             =   ""
                  Object.Tag             =   ""
                  ImageVarType    =   2
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame frPayrollItems 
         Height          =   6255
         Left            =   120
         TabIndex        =   129
         Top             =   120
         Visible         =   0   'False
         Width           =   9495
         Begin VB.CommandButton cmdPyrlRefresh 
            Caption         =   "Refresh"
            Enabled         =   0   'False
            Height          =   375
            Left            =   7680
            TabIndex        =   167
            Top             =   660
            Width           =   1695
         End
         Begin VB.CommandButton cmdPyrlUpdate 
            Caption         =   "Update"
            Enabled         =   0   'False
            Height          =   375
            Left            =   7680
            TabIndex        =   166
            Top             =   240
            Width           =   1695
         End
         Begin VB.CheckBox chkPyrllItem 
            Caption         =   "Employer"
            DataField       =   "EmployerYN"
            Height          =   255
            Left            =   240
            TabIndex        =   165
            Top             =   3960
            Width           =   975
         End
         Begin VB.TextBox txtPyrllItems 
            DataField       =   "PyrlItemID"
            Height          =   285
            Index           =   0
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   164
            Top             =   600
            Width           =   1695
         End
         Begin VB.ComboBox cbPyrllItems 
            DataField       =   "Type"
            Height          =   315
            Index           =   0
            ItemData        =   "frm_Sys_Setup_Employee.frx":0E2A
            Left            =   1320
            List            =   "frm_Sys_Setup_Employee.frx":0E3A
            TabIndex        =   163
            Top             =   960
            Width           =   2055
         End
         Begin VB.ComboBox cbPyrllItems 
            DataField       =   "Basis"
            Height          =   315
            Index           =   1
            ItemData        =   "frm_Sys_Setup_Employee.frx":0E69
            Left            =   1320
            List            =   "frm_Sys_Setup_Employee.frx":0E76
            TabIndex        =   162
            Top             =   1320
            Width           =   2055
         End
         Begin VB.TextBox txtPyrllItems 
            DataField       =   "Description"
            Height          =   285
            Index           =   1
            Left            =   5280
            TabIndex        =   161
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox txtPyrllItems 
            DataField       =   "WageLow"
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
            Index           =   2
            Left            =   5280
            TabIndex        =   160
            Top             =   960
            Width           =   1695
         End
         Begin VB.TextBox txtPyrllItems 
            DataField       =   "WageHigh"
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
            Index           =   3
            Left            =   5280
            TabIndex        =   159
            Top             =   1320
            Width           =   1695
         End
         Begin VB.TextBox txtPyrllItems 
            DataField       =   "YTDMax"
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
            Index           =   4
            Left            =   5280
            TabIndex        =   158
            Top             =   1680
            Width           =   1695
         End
         Begin VB.TextBox txtPyrllItems 
            Height          =   285
            Index           =   5
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   157
            Top             =   1680
            Width           =   2055
         End
         Begin VB.ComboBox cbPyrllItems 
            DataField       =   "Default"
            Height          =   315
            Index           =   2
            ItemData        =   "frm_Sys_Setup_Employee.frx":0E8B
            Left            =   7680
            List            =   "frm_Sys_Setup_Employee.frx":0EC2
            TabIndex        =   156
            Top             =   1800
            Width           =   1695
         End
         Begin VB.Frame frEmployee 
            Caption         =   "Employee"
            Height          =   1695
            Left            =   120
            TabIndex        =   145
            Top             =   2160
            Width           =   9255
            Begin VB.TextBox txtFieldsTemp 
               DataField       =   " "
               DataSource      =   "adoPrimaryRS"
               Height          =   285
               Left            =   7200
               Locked          =   -1  'True
               TabIndex        =   150
               Text            =   "231100"
               Top             =   1320
               Visible         =   0   'False
               Width           =   1935
            End
            Begin VB.TextBox txtPyrllItems 
               DataField       =   "ItemAmount"
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
               Index           =   6
               Left            =   1680
               TabIndex        =   149
               Top             =   480
               Width           =   1695
            End
            Begin VB.TextBox txtPyrllItems 
               DataField       =   "ItemPercent"
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
               Index           =   7
               Left            =   1680
               TabIndex        =   148
               Top             =   960
               Width           =   1695
            End
            Begin VB.TextBox txtPyrllItems 
               DataField       =   "Account"
               Height          =   285
               Index           =   20
               Left            =   5160
               Locked          =   -1  'True
               TabIndex        =   147
               Top             =   480
               Width           =   1695
            End
            Begin VB.CommandButton cmdAcct 
               Height          =   270
               Index           =   20
               Left            =   6840
               Picture         =   "frm_Sys_Setup_Employee.frx":0FEC
               Style           =   1  'Graphical
               TabIndex        =   146
               Top             =   480
               Width           =   375
            End
            Begin VB.Label lblAcctTemp 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   7200
               TabIndex        =   155
               Top             =   1080
               Visible         =   0   'False
               Width           =   1935
            End
            Begin VB.Label lblPyrllItems 
               Caption         =   "Amount:"
               Height          =   255
               Index           =   9
               Left            =   240
               TabIndex        =   154
               Top             =   480
               Width           =   1335
            End
            Begin VB.Label lblPyrllItems 
               Caption         =   "Percent of Basis:"
               Height          =   255
               Index           =   10
               Left            =   240
               TabIndex        =   153
               Top             =   960
               Width           =   1335
            End
            Begin VB.Label lblPyrllItems 
               Caption         =   "Account:"
               Height          =   255
               Index           =   20
               Left            =   3840
               TabIndex        =   152
               Top             =   480
               Width           =   1335
            End
            Begin VB.Label lblAccts 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   20
               Left            =   3840
               TabIndex        =   151
               Top             =   780
               Width           =   3375
            End
         End
         Begin VB.CommandButton cmdItemID 
            Height          =   270
            Left            =   3000
            Picture         =   "frm_Sys_Setup_Employee.frx":1136
            Style           =   1  'Graphical
            TabIndex        =   144
            Top             =   600
            Width           =   375
         End
         Begin VB.Frame frAccount 
            Height          =   2175
            Left            =   120
            TabIndex        =   131
            Top             =   3960
            Width           =   9255
            Begin VB.TextBox txtPyrllItems 
               DataField       =   "EmployerItemAmount"
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
               Left            =   1680
               TabIndex        =   137
               Top             =   480
               Width           =   1695
            End
            Begin VB.TextBox txtPyrllItems 
               DataField       =   "EmployerItemPercent"
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
               Index           =   10
               Left            =   1680
               TabIndex        =   136
               Top             =   960
               Width           =   1695
            End
            Begin VB.TextBox txtPyrllItems 
               DataField       =   "Account2"
               Height          =   285
               Index           =   21
               Left            =   5280
               Locked          =   -1  'True
               TabIndex        =   135
               Top             =   480
               Width           =   1695
            End
            Begin VB.TextBox txtPyrllItems 
               DataField       =   "Account3"
               Height          =   285
               Index           =   22
               Left            =   5280
               Locked          =   -1  'True
               TabIndex        =   134
               Top             =   1320
               Width           =   1695
            End
            Begin VB.CommandButton cmdAcct 
               Height          =   270
               Index           =   22
               Left            =   6960
               Picture         =   "frm_Sys_Setup_Employee.frx":1280
               Style           =   1  'Graphical
               TabIndex        =   133
               Top             =   1320
               Width           =   375
            End
            Begin VB.CommandButton cmdAcct 
               Height          =   270
               Index           =   21
               Left            =   6960
               Picture         =   "frm_Sys_Setup_Employee.frx":13CA
               Style           =   1  'Graphical
               TabIndex        =   132
               Top             =   480
               Width           =   375
            End
            Begin VB.Label lblPyrllItems 
               Caption         =   "Amount:"
               Height          =   255
               Index           =   12
               Left            =   240
               TabIndex        =   143
               Top             =   480
               Width           =   1335
            End
            Begin VB.Label lblPyrllItems 
               Caption         =   "Percent of Basis:"
               Height          =   255
               Index           =   13
               Left            =   240
               TabIndex        =   142
               Top             =   960
               Width           =   1335
            End
            Begin VB.Label lblPyrllItems 
               Caption         =   "Debit Account:"
               Height          =   255
               Index           =   21
               Left            =   3840
               TabIndex        =   141
               Top             =   480
               Width           =   1335
            End
            Begin VB.Label lblPyrllItems 
               Caption         =   "Credit Account:"
               Height          =   255
               Index           =   22
               Left            =   3840
               TabIndex        =   140
               Top             =   1320
               Width           =   1335
            End
            Begin VB.Label lblAccts 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   21
               Left            =   3840
               TabIndex        =   139
               Top             =   780
               Width           =   3495
            End
            Begin VB.Label lblAccts 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   22
               Left            =   3840
               TabIndex        =   138
               Top             =   1620
               Width           =   3495
            End
         End
         Begin VB.CommandButton cmdClosePyrlItem 
            Caption         =   "Back"
            Height          =   375
            Left            =   7680
            TabIndex        =   130
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Label lblPyrllItems 
            Caption         =   "Item ID:"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   176
            Top             =   600
            Width           =   975
         End
         Begin VB.Label lblPyrllItems 
            Caption         =   "Type:"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   175
            Top             =   960
            Width           =   975
         End
         Begin VB.Label lblPyrllItems 
            Caption         =   "Basis:"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   174
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label lblPyrllItems 
            Caption         =   "Description:"
            Height          =   255
            Index           =   3
            Left            =   3840
            TabIndex        =   173
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label lblPyrllItems 
            Caption         =   "YTD Gross Low:"
            Height          =   255
            Index           =   4
            Left            =   3840
            TabIndex        =   172
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label lblPyrllItems 
            Caption         =   "YTD Gross High:"
            Height          =   255
            Index           =   5
            Left            =   3840
            TabIndex        =   171
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label lblPyrllItems 
            Caption         =   "Maximum Annual:"
            Height          =   255
            Index           =   6
            Left            =   3840
            TabIndex        =   170
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label lblPyrllItems 
            Caption         =   "Desc:"
            Height          =   255
            Index           =   7
            Left            =   240
            TabIndex        =   169
            Top             =   1680
            Width           =   975
         End
         Begin VB.Label lblPyrllItems 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Item Tracking:"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   8
            Left            =   7680
            TabIndex        =   168
            Top             =   1560
            Width           =   1695
         End
      End
      Begin VB.Frame frBeginningBal 
         Height          =   7480
         Left            =   120
         TabIndex        =   79
         Top             =   120
         Visible         =   0   'False
         Width           =   10935
         Begin VB.Frame Frame4 
            Enabled         =   0   'False
            Height          =   1455
            Left            =   3480
            TabIndex        =   111
            Top             =   720
            Width           =   4815
            Begin VB.TextBox txt 
               Height          =   285
               Index           =   0
               Left            =   1560
               Locked          =   -1  'True
               TabIndex        =   114
               Top             =   600
               Width           =   2535
            End
            Begin VB.ComboBox cboPeriod 
               Height          =   315
               ItemData        =   "frm_Sys_Setup_Employee.frx":1514
               Left            =   1560
               List            =   "frm_Sys_Setup_Employee.frx":1524
               TabIndex        =   113
               Top             =   960
               Width           =   1695
            End
            Begin VB.TextBox txtNew 
               Height          =   285
               Index           =   0
               Left            =   1560
               Locked          =   -1  'True
               TabIndex        =   112
               Top             =   240
               Width           =   1335
            End
            Begin VB.Label lblBeginning 
               Caption         =   "Employee:"
               Height          =   255
               Index           =   0
               Left            =   720
               TabIndex        =   117
               Top             =   600
               Width           =   735
            End
            Begin VB.Label lblBeginning 
               Alignment       =   1  'Right Justify
               Caption         =   "Period:"
               Height          =   255
               Index           =   9
               Left            =   720
               TabIndex        =   116
               Top             =   960
               Width           =   735
            End
            Begin VB.Label lblBeginning 
               Alignment       =   1  'Right Justify
               Caption         =   "Employee ID:"
               Height          =   255
               Index           =   10
               Left            =   480
               TabIndex        =   115
               Top             =   240
               Width           =   975
            End
         End
         Begin VB.Frame Frame3 
            Height          =   1455
            Left            =   120
            TabIndex        =   104
            Top             =   720
            Width           =   3255
            Begin VB.CommandButton cmdDate 
               Height          =   285
               Index           =   1
               Left            =   2520
               Picture         =   "frm_Sys_Setup_Employee.frx":1569
               Style           =   1  'Graphical
               TabIndex        =   108
               Top             =   480
               Width           =   375
            End
            Begin VB.TextBox txt 
               Height          =   285
               Index           =   2
               Left            =   1200
               Locked          =   -1  'True
               TabIndex        =   107
               Top             =   960
               Width           =   1335
            End
            Begin VB.CommandButton cmdDate 
               Height          =   285
               Index           =   2
               Left            =   2520
               Picture         =   "frm_Sys_Setup_Employee.frx":1B43
               Style           =   1  'Graphical
               TabIndex        =   106
               Top             =   960
               Width           =   375
            End
            Begin VB.TextBox txt 
               Height          =   285
               Index           =   1
               Left            =   1200
               Locked          =   -1  'True
               TabIndex        =   105
               Top             =   480
               Width           =   1335
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               Caption         =   "EndDate:"
               Height          =   255
               Left            =   240
               TabIndex        =   110
               Top             =   960
               Width           =   855
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               Caption         =   "StartDate:"
               Height          =   255
               Left            =   240
               TabIndex        =   109
               Top             =   480
               Width           =   855
            End
         End
         Begin VB.TextBox txtBeginning 
            DataField       =   "GROSS"
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
            Left            =   120
            TabIndex        =   103
            Top             =   3240
            Width           =   1335
         End
         Begin VB.TextBox txtBeginning 
            DataField       =   "FICA"
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
            Index           =   2
            Left            =   1440
            TabIndex        =   102
            Top             =   3240
            Width           =   1335
         End
         Begin VB.TextBox txtBeginning 
            DataField       =   "FIT"
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
            Index           =   3
            Left            =   2760
            TabIndex        =   101
            Top             =   3240
            Width           =   1335
         End
         Begin VB.TextBox txtBeginning 
            DataField       =   "STATETAX"
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
            Index           =   4
            Left            =   4080
            TabIndex        =   100
            Top             =   3240
            Width           =   1335
         End
         Begin VB.TextBox txtBeginning 
            DataField       =   "LOCAL"
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
            Index           =   5
            Left            =   5400
            TabIndex        =   99
            Top             =   3240
            Width           =   1335
         End
         Begin VB.TextBox txtBeginning 
            DataField       =   "REGHOURS"
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
            Index           =   6
            Left            =   6720
            TabIndex        =   98
            Top             =   3240
            Width           =   1335
         End
         Begin VB.TextBox txtBeginning 
            DataField       =   "OTHOURS"
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
            Index           =   7
            Left            =   8040
            TabIndex        =   97
            Top             =   3240
            Width           =   1335
         End
         Begin VB.TextBox txtOld 
            Height          =   285
            Index           =   7
            Left            =   8040
            TabIndex        =   96
            Top             =   2520
            Width           =   1335
         End
         Begin VB.TextBox txtOld 
            Height          =   285
            Index           =   6
            Left            =   6720
            TabIndex        =   95
            Top             =   2520
            Width           =   1335
         End
         Begin VB.TextBox txtOld 
            Height          =   285
            Index           =   5
            Left            =   5400
            TabIndex        =   94
            Top             =   2520
            Width           =   1335
         End
         Begin VB.TextBox txtOld 
            Height          =   285
            Index           =   4
            Left            =   4080
            TabIndex        =   93
            Top             =   2520
            Width           =   1335
         End
         Begin VB.TextBox txtOld 
            Height          =   285
            Index           =   3
            Left            =   2760
            TabIndex        =   92
            Top             =   2520
            Width           =   1335
         End
         Begin VB.TextBox txtOld 
            Height          =   285
            Index           =   2
            Left            =   1440
            TabIndex        =   91
            Top             =   2520
            Width           =   1335
         End
         Begin VB.TextBox txtOld 
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   90
            Top             =   2520
            Width           =   1335
         End
         Begin VB.TextBox txtNew 
            Height          =   285
            Index           =   7
            Left            =   8040
            TabIndex        =   89
            Top             =   2880
            Width           =   1335
         End
         Begin VB.TextBox txtNew 
            Height          =   285
            Index           =   6
            Left            =   6720
            TabIndex        =   88
            Top             =   2880
            Width           =   1335
         End
         Begin VB.TextBox txtNew 
            Height          =   285
            Index           =   5
            Left            =   5400
            TabIndex        =   87
            Top             =   2880
            Width           =   1335
         End
         Begin VB.TextBox txtNew 
            Height          =   285
            Index           =   4
            Left            =   4080
            TabIndex        =   86
            Top             =   2880
            Width           =   1335
         End
         Begin VB.TextBox txtNew 
            Height          =   285
            Index           =   3
            Left            =   2760
            TabIndex        =   85
            Top             =   2880
            Width           =   1335
         End
         Begin VB.TextBox txtNew 
            Height          =   285
            Index           =   2
            Left            =   1440
            TabIndex        =   84
            Top             =   2880
            Width           =   1335
         End
         Begin VB.TextBox txtNew 
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   83
            Top             =   2880
            Width           =   1335
         End
         Begin VB.TextBox txtBeginning 
            Alignment       =   2  'Center
            DataField       =   "PRINTED"
            BeginProperty DataFormat 
               Type            =   5
               Format          =   """$""#,##0.00;(""$""#,##0.00)"
               HaveTrueFalseNull=   1
               TrueValue       =   "Posted"
               FalseValue      =   "Not Posted"
               NullValue       =   "New Data"
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   7
            EndProperty
            Height          =   285
            Index           =   8
            Left            =   9360
            Locked          =   -1  'True
            TabIndex        =   82
            Top             =   2520
            Width           =   1335
         End
         Begin VB.CommandButton cmdClose 
            Caption         =   "Back"
            Height          =   855
            Index           =   1
            Left            =   8640
            Picture         =   "frm_Sys_Setup_Employee.frx":211D
            Style           =   1  'Graphical
            TabIndex        =   81
            Top             =   1080
            Width           =   975
         End
         Begin VB.CommandButton cmdPost 
            Caption         =   "Post"
            Height          =   855
            Index           =   0
            Left            =   9720
            Picture         =   "frm_Sys_Setup_Employee.frx":2427
            Style           =   1  'Graphical
            TabIndex        =   80
            Top             =   1080
            Width           =   975
         End
         Begin MSDataGridLib.DataGrid DataGrid2 
            Height          =   2655
            Left            =   5520
            TabIndex        =   118
            Top             =   3720
            Width           =   5295
            _ExtentX        =   9340
            _ExtentY        =   4683
            _Version        =   393216
            AllowUpdate     =   -1  'True
            HeadLines       =   1
            RowHeight       =   15
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
            Caption         =   "Posted Items"
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
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
               DataField       =   ""
               Caption         =   ""
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
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Bindings        =   "frm_Sys_Setup_Employee.frx":2869
            Height          =   2655
            Left            =   120
            TabIndex        =   119
            Top             =   3720
            Width           =   5295
            _ExtentX        =   9340
            _ExtentY        =   4683
            _Version        =   393216
            AllowUpdate     =   -1  'True
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
            Caption         =   "Items not posted"
            ColumnCount     =   3
            BeginProperty Column00 
               DataField       =   "PyrlItemID"
               Caption         =   "Item ID"
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
               DataField       =   "Description"
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
            BeginProperty Column02 
               DataField       =   "TotalAmount"
               Caption         =   "Total Amount"
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
                  ColumnWidth     =   1140.095
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1844.787
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   1739.906
               EndProperty
            EndProperty
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000001&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Beginning Balances By Quarter"
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   2
            Left            =   80
            TabIndex        =   128
            Top             =   240
            Width           =   10785
         End
         Begin VB.Label lblBeginning 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Posted"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   8
            Left            =   9360
            TabIndex        =   127
            Top             =   2280
            Width           =   1335
         End
         Begin VB.Label lblBeginning 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "OT Hours"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   7
            Left            =   8040
            TabIndex        =   126
            Top             =   2280
            Width           =   1335
         End
         Begin VB.Label lblBeginning 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Reg Hours"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   6
            Left            =   6720
            TabIndex        =   125
            Top             =   2280
            Width           =   1335
         End
         Begin VB.Label lblBeginning 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Local Tax"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   5
            Left            =   5400
            TabIndex        =   124
            Top             =   2280
            Width           =   1335
         End
         Begin VB.Label lblBeginning 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "State Tax"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   4
            Left            =   4080
            TabIndex        =   123
            Top             =   2280
            Width           =   1335
         End
         Begin VB.Label lblBeginning 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "FIT"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   2760
            TabIndex        =   122
            Top             =   2280
            Width           =   1335
         End
         Begin VB.Label lblBeginning 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "FICA"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   1440
            TabIndex        =   121
            Top             =   2280
            Width           =   1335
         End
         Begin VB.Label lblBeginning 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "GROSS"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   120
            Top             =   2280
            Width           =   1335
         End
      End
   End
   Begin VB.Label lblTop 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Employee Setup"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   194
      Top             =   120
      Width           =   10785
   End
End
Attribute VB_Name = "frm_SYS_Setup_Employee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ADOprimaryrs As ADODB.Recordset
Attribute ADOprimaryrs.VB_VarHelpID = -1
Dim ADOemployee As ADODB.Recordset
Dim ADOCreatePay As ADODB.Recordset

Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean

Dim db As ADODB.Connection
'The recordset should only contain one record holding information pertaining to  a specific company
' inventory setup.

Private Function CbValidate(cbName As ComboBox, cbText As String) As Boolean
Dim i As Integer
    For i = 0 To cbName.ListCount - 1
        If cbText = cbName.List(i) Then
            CbValidate = True
            Exit Function
        End If
    Next
    cbName.Text = cbName.List(0)
    CbValidate = False
End Function

Private Sub cbAE_Click(Index As Integer)
Select Case Index
Case 2
    PayTypeSelection
End Select

End Sub

Private Sub PayTypeSelection()
    If cbAE(2).Text = "" Then
        MsgBox "Please select Pay Type first"
        Exit Sub
    Else
                txtAE(0).Visible = False
                lblAE(4).Visible = False
                picCommision.Visible = False
                txtAE(2).Visible = False
                lblAE(6).Visible = False
                txtAE(3).Visible = False
                lblAE(7).Visible = False
        cbAE(3).Clear
        If cbAE(2).Text = "Hourly" Then
                txtAE(2).Visible = True
                lblAE(6).Visible = True
                txtAE(3).Visible = True
                lblAE(7).Visible = True
            cbAE(3).List(0) = "Weekly"
            cbAE(3).List(1) = "Biweekly"
                If ADOemployee Is Nothing Then
                Else
                    ADOemployee![EMP Salesperson YN] = False
                End If
        Else
            If cbAE(2).Text = "Salary" Then
                txtAE(0).Visible = True
                lblAE(4).Visible = True
                If ADOemployee Is Nothing Then
                Else
                    ADOemployee![EMP Salesperson YN] = False
                End If
            ElseIf cbAE(2).Text = "Salary + Commission" Then  '-
                If ADOemployee Is Nothing Then
                Else
                    ADOemployee![EMP Salesperson YN] = True
                End If
                txtAE(0).Visible = True
                lblAE(4).Visible = True
                picCommision.Visible = True
            ElseIf cbAE(2).Text = "Commission" Then  '-
                picCommision.Visible = True
                If ADOemployee Is Nothing Then
                Else
                    ADOemployee![EMP Salesperson YN] = True
                End If
            End If
            
            cbAE(3).List(0) = "Weekly"
            cbAE(3).List(1) = "Biweekly"
            cbAE(3).List(2) = "SemiMonthly"
            cbAE(3).List(3) = "Monthly"
        End If
    End If
End Sub

Private Sub cbAE_KeyPress(Index As Integer, KeyAscii As Integer)
    keyResponse = CtrlValidate(KeyAscii, "")
    If keyResponse = True Then
    Else
       KeyAscii = 0
    End If
End Sub

Private Sub cbAE_LostFocus(Index As Integer)
    If CbValidate(cbAE(Index), cbAE(Index).Text) = False Then
       MsgBox "There is no such selection", vbInformation, "Information"
    End If
End Sub

Private Sub cbDT_KeyPress(KeyAscii As Integer)
    keyResponse = CtrlValidate(KeyAscii, "")
    If keyResponse = True Then
    Else
       KeyAscii = 0
    End If
End Sub

Private Sub cbDT_LostFocus()
    If CbValidate(cbDT, cbDT.Text) = False Then
       MsgBox "There is no such selection", vbInformation, "Information"
    End If
End Sub

Private Sub cboPeriod_Click()
Dim ThisYear As String
Dim StartDate
Dim EndDate

ThisYear = Format(Date, "yyyy")

Select Case cboPeriod.ListIndex
Case 0
    ThisYear = "1/1/" & ThisYear
    StartDate = FormatDate(CDate(ThisYear))
    ThisYear = Format(Date, "yyyy")
    ThisYear = "3/31/" & ThisYear
    EndDate = FormatDate(CDate(ThisYear))
Case 1
    ThisYear = "4/1/" & ThisYear
    StartDate = FormatDate(CDate(ThisYear))
    ThisYear = Format(Date, "yyyy")
    ThisYear = "6/30/" & ThisYear
    EndDate = FormatDate(CDate(ThisYear))
Case 2
    ThisYear = "7/1/" & ThisYear
    StartDate = FormatDate(CDate(ThisYear))
    ThisYear = Format(Date, "yyyy")
    ThisYear = "9/30/" & ThisYear
    EndDate = FormatDate(CDate(ThisYear))
Case 3
    ThisYear = "10/1/" & ThisYear
    StartDate = FormatDate(CDate(ThisYear))
    ThisYear = Format(Date, "yyyy")
    ThisYear = "12/31/" & ThisYear
    EndDate = FormatDate(CDate(ThisYear))
End Select

db.Execute "Delete * from [Pyrl - ItemsBegBalWork]"

'- code below is equal to ---->> DoCmd.OpenQuery ("Pyrl - ItemBegBalUnmatched")
db.Execute "INSERT INTO [Pyrl - ItemsBegBalWork] ( ItemID, Description, TotalAmount, Basis, Type )" & _
           " SELECT DISTINCTROW [Pyrl - ItemBegBalItems].PyrlItemID, [Pyrl - ItemBegBalItems].Description, [Pyrl - ItemBegBalItems].TotalAmount, [Pyrl - ItemBegBalItems].Basis, [Pyrl - ItemBegBalItems].Type" & _
           " FROM [Pyrl - ItemBegBalItems] LEFT JOIN [Pyrl - ItemsBegBalRegister] ON [Pyrl - ItemBegBalItems].PyrlItemID = [Pyrl - ItemsBegBalRegister].PyrlItemID" & _
           " Where ((([Pyrl - ItemsBegBalRegister].PyrlItemID) Is Null))"
'- code below is equal to ---->> DoCmd.OpenQuery ("Pyrl - ItemBegBalW")
db.Execute "UPDATE [Pyrl - ItemsBegBalWork] SET [Pyrl - ItemsBegBalWork].EmployeeID ='" & txtNew(0).Text & "', [Pyrl - ItemsBegBalWork].EmployeeName ='" & txtBeginning(0) & "'"
        
Dim i As Integer
    For i = 1 To txtNew.UBound
                txtNew(i) = 0
    Next
    If txtBeginning(8) = "Posted" Then
        For i = 1 To txtOld.UBound
            txtOld(i) = txtBeginning(1)
        Next
    Else
        For i = 1 To txtOld.UBound
            txtOld(i) = 0
        Next
    End If
End Sub

Private Sub cboPeriod_KeyPress(KeyAscii As Integer)
    keyResponse = CtrlValidate(KeyAscii, "")
    If keyResponse = True Then
    Else
       KeyAscii = 0
    End If
End Sub

Private Sub cboPeriod_LostFocus()
    If CbValidate(cboPeriod, cboPeriod.Text) = False Then
       MsgBox "There is no such selection", vbInformation, "Information"
    End If
End Sub

Private Sub cbPyrllItems_KeyPress(Index As Integer, KeyAscii As Integer)
    keyResponse = CtrlValidate(KeyAscii, "")
    If keyResponse = True Then
    Else
       KeyAscii = 0
    End If
End Sub

Private Sub cbPyrllItems_LostFocus(Index As Integer)
    If CbValidate(cbPyrllItems(Index), cbPyrllItems(Index).Text) = False Then
       MsgBox "There is no such selection", vbInformation, "Information"
    End If
    TypeDesc
End Sub


Private Sub chkPyrllItem_Click()
    If chkPyrllItem.Value = 1 Then
        frAccount.Enabled = True
    Else
        frAccount.Enabled = False
    End If
End Sub

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
    lblAccts(Index).Caption = lblAcctTemp.Caption
    txtPyrllItems(Index).Text = txtFieldsTemp.Text
    txtPyrllItems(Index).SetFocus
End Sub

Private Sub cmdAdd_Click()
  On Error GoTo cmdNew_Click_Error
  
  'Call UpdatePyrlItems
  'Call ResetPyrlItems
  'Me.Refresh
  'DoCmd.GoToRecord acForm, Me.Name, acNewRec
  'Me![EMP ID].SetFocus
  
  'Me!cmdDelete.Enabled = True
  
  
  Exit Sub
cmdNew_Click_Error:
  Call ErrorLog("Setup Employee", "cmdAdd_Click", Now, Err.Number, Err.Description, True, db)
  Resume Next
End Sub

Private Sub cmdClose_Click(Index As Integer)
Select Case Index
Case 0
    Unload Me
Case 1
    frPayrollItems.Visible = False
    frBeginningBal.Visible = False
    frEmployeeSetup.Visible = True
    Me.Caption = "Employee Setup"
    picButtons.Visible = True
    Form_Resize
End Select
End Sub

Private Sub cmdClosePyrlItem_Click()

    If txtPyrllItems(0) <> "" Then         '-enable this when finish
        ADOCreatePay.CancelUpdate
        ADOCreatePay.Close
    End If
    
    Set ADOCreatePay = Nothing
    
    Dim EmptyText As TextBox
    For Each EmptyText In Me.txtPyrllItems
        Set EmptyText.DataSource = Nothing
        EmptyText = ""
    Next
    Dim Emptycb As ComboBox
    For Each Emptycb In Me.cbPyrllItems
        Set Emptycb.DataSource = Nothing
        Emptycb = ""
    Next
    'frPayrollItems.Top = 0
    'frPayrollItems.Left = 0
    Me.Caption = "Employee Setup"
    frEmployeeSetup.Visible = True
    frBeginningBal.Visible = False
    frPayrollItems.Visible = False
    Form_Resize
    cmdPyrlUpdate.Enabled = False
    cmdPyrlRefresh.Enabled = False
    chkPyrllItem.Enabled = False
End Sub

Private Sub cmdCreatePyrll_Click()
'Dim ctrl
    frEmployeeSetup.Visible = False
    frBeginningBal.Visible = False
    frPayrollItems.ZOrder 0
    frPayrollItems.Visible = True
    Me.Caption = "Employee Setup - Create Payroll"
    Form_Resize
    frEmployee.Enabled = False
    frAccount.Enabled = False
    chkPyrllItem.Enabled = False
End Sub

Private Sub cmdDate_Click(Index As Integer)
Select Case Index
Case 0
    Menu_Calendar.WhoCallMe True, 1555
    'Menu_Calendar.Show vbModal
Case 1
    Menu_Calendar.WhoCallMe True, 1620
    'Menu_Calendar.Show vbModal
    LoadbeginBal
Case 2
    Menu_Calendar.WhoCallMe True, 1630
    'Menu_Calendar.Show vbModal
    LoadbeginBal
Case 3
    Menu_Calendar.WhoCallMe True, 1556
End Select
End Sub

Private Sub LoadbeginBal()

MsgBox "I do think this wrong, Please check the Pyrl-Register and Pyrl-Register Detail"
Exit Sub

Dim cmdItem As Command
Dim rsItem As ADODB.Recordset
Dim ParamItem1 As Parameter
Dim ParamItem2 As Parameter
Dim ParamItem3 As Parameter
    
If txt(1) <> "" And txt(2) <> "" Then
    Frame4.Enabled = True
    
    Set cmdItem = New Command
    cmdItem.ActiveConnection = db
        
    cmdItem.CommandText = "[Pyrl - ItemsBegBalRegister]"         'Pyrl - ItemsBegBalRegister
    cmdItem.CommandType = adCmdStoredProc
        
    Set ParamItem1 = cmdItem.CreateParameter("[Forms]![Employee Card]![Emp ID]", adBSTR, adParamInput) 'Screen.ActiveForm.[AP PAY Check No]       'set query criteria for current work table records
    Set ParamItem2 = cmdItem.CreateParameter("[Forms]![Pyrl - BeginningBalances]![StartDate]", adDate, adParamInput) 'Screen.ActiveForm.[AP PAY Check No]       'set query criteria for current work table records
    Set ParamItem3 = cmdItem.CreateParameter("[Forms]![Pyrl - BeginningBalances]![EndDate]", adDate, adParamInput) 'Screen.ActiveForm.[AP PAY Check No]       'set query criteria for current work table records
        
    ParamItem1.Value = txtfields(1)
    ParamItem2.Value = txt(1)
    ParamItem3.Value = txt(2)
        
    cmdItem.Parameters.Append ParamItem1
    cmdItem.Parameters.Append ParamItem2
    cmdItem.Parameters.Append ParamItem3
        
    Set rsItem = cmdItem.Execute
        
    Set DataGrid1.DataSource = rsItem
End If
End Sub

Private Sub cmdDelete_Click()
  On Error GoTo cmdDelete_Click_Error
  
  'DeleteRec.Value = "yes"
  'Me.Refresh
    
  'If NZ(Me![EMP ID], "") <> "" Then
  '      Dim rsEmp As Recordset
  '      Set rsEmp = db.OpenRecordset("Select * FROM [Pyrl - Register] Where [EMP ID] = '" & Me![EMP ID] & "' AND [Printed] = 0")
  '      If rsEmp.RecordCount > 0 Then
  '          MsgBox ("This employee has " & rsEmp.RecordCount & " unposted payroll check(s) in the " & Chr(10) & "payroll register. Either print the checks or delete the created " & Chr(10) & "payroll entry thru the Pay Employees 'Preview' screen.")
  '          Exit Sub
  '      End If
        
  '      Call ResetPyrlItems
  '      DoCmd.RunMacro "Delete Record"
  '      DoCmd.GoToRecord acForm, Me.Name, acNewRec
  '      Me![EMP ID].Value = ""
  '      Me![EMP ID].SetFocus
  'End If
  
  'DeleteRec.Value = "No"
  'Me.Refresh

Exit Sub
cmdDelete_Click_Error:
  Call ErrorLog("Setup Employee", "cmdDelete_Click", Now, Err.Number, Err.Description, True, db)
  Resume Next
End Sub

Private Sub cmdEmpID_Click()
    Dim SQLstatement As String
    Dim No As Integer
    Dim ghead As String
    Dim fhead As String

    No = 1530
    SQLstatement = "select [EMP ID], [EMP Name]" & _
                    "from [EMP Employees]"
    ghead = "Employee"
    fhead = "ID//Name"
    
    AllLookup.GetWhichTable No, SQLstatement, ghead, fhead, db
    'AllLookup.Show vbModal
    OpenDBemp
    PayTypeSelection
End Sub

Private Sub OpenDBemp()
ShowStatus True
Dim TempName As String

Dim Ctrl As Control
Dim oText As TextBox
Dim cb As ComboBox
Dim chk As CheckBox
    
    If txtfields(1) <> "" Then
        Set ADOemployee = New ADODB.Recordset
        ADOemployee.Open "Select * FROM [Pyrl - Employees] Where [EMP ID] = '" & txtfields(1) & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
        
        If ADOemployee.RecordCount = 0 Then
            Dim iResponse As Integer
            iResponse = MsgBox("There is no such employee ID. Would you like create a new Employee", vbYesNo, "Employee")
            If iResponse = vbNo Then
                For Each Ctrl In Me.Controls
                    If TypeOf Ctrl Is TextBox Or TypeOf Ctrl Is ComboBox Then
                       Set Ctrl.DataSource = Nothing
                       Ctrl.Text = ""
                    End If
                Next
                
                ADOemployee.CancelUpdate
                ADOemployee.Close
                Set ADOemployee = Nothing
                
                ShowStatus False
                cmdUpdate.Enabled = False
                cmdRefresh.Enabled = False
                Exit Sub
            Else
                TempName = txtfields(1)
                ADOemployee.AddNew
            End If
        End If
            
            cmdUpdate.Enabled = True
            cmdRefresh.Enabled = True
            cmdCreatePyrll.Enabled = True
            cmdBeginning.Enabled = True
            
            'Bind the text boxes to the data provider cbAE txtAE chkAE txtDT cbDT
            For Each chk In Me.chkAE
              Set chk.DataSource = ADOemployee
            Next
            
            Set cbDT.DataSource = ADOemployee
            
            For Each cb In Me.cbAE
              Set cb.DataSource = ADOemployee
            Next
            For Each oText In Me.txtfields
              If Trim(oText.DataField) <> "" Then
                Set oText.DataSource = ADOemployee
                If ADOemployee("" & oText.DataField & "").Type = 202 Then oText.MaxLength = ADOemployee("" & oText.DataField & "").DefinedSize
              End If
            Next
            For Each oText In Me.txtAE
              Set oText.DataSource = ADOemployee
              If ADOemployee("" & oText.DataField & "").Type = 202 Then oText.MaxLength = ADOemployee("" & oText.DataField & "").DefinedSize
            Next
            For Each oText In Me.txtDT
              Set oText.DataSource = ADOemployee
              If ADOemployee("" & oText.DataField & "").Type = 202 Then oText.MaxLength = ADOemployee("" & oText.DataField & "").DefinedSize
            Next
            If ADOemployee.EditMode = adEditAdd Then
                txtfields(1) = TempName
            End If
    End If
ShowStatus False
End Sub

Private Sub cmdItemID_Click()
    Dim SQLstatement As String
    Dim No As Integer
    Dim ghead As String
    Dim fhead As String

    No = 1610
    SQLstatement = "select [PyrlItemID], [Description]" & _
                    "from [Pyrl - Payroll Items]"
    ghead = "Payroll Items"
    fhead = "ID//Description"
    
    AllLookup.GetWhichTable No, SQLstatement, ghead, fhead, db
    'AllLookup.Show vbModal
    txtPyrllItems(0).SetFocus
    If txtPyrllItems(0) <> "" Then
    
    Set ADOCreatePay = New ADODB.Recordset
    ADOCreatePay.Open "SELECT * FROM [Pyrl - Payroll Items] WHERE [PyrlItemID]='" & txtPyrllItems(0) & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
    Dim textPay As TextBox
    For Each textPay In Me.txtPyrllItems
        Set textPay.DataSource = ADOCreatePay
        Select Case textPay.Index
        Case 20, 21, 22
           If textPay.Text <> "" Then
                lblAccts(textPay.Index) = LookRecord("[GL COA Account Name]", "[GL Chart Of Accounts]", db, "[GL COA Account No] = '" & textPay.Text & "'")
           End If
        End Select
    Next
    
    Set chkPyrllItem.DataSource = ADOCreatePay
    
    Dim cbPay As ComboBox
    For Each cbPay In Me.cbPyrllItems
        Set cbPay.DataSource = ADOCreatePay
    Next
    
    frEmployee.Enabled = True
    If ADOCreatePay![EmployerYN] = True Then
        frAccount.Enabled = True
    Else
        frAccount.Enabled = False
    End If
    
        TypeDesc
    End If
    cmdPyrlUpdate.Enabled = True
    cmdPyrlRefresh.Enabled = True
    chkPyrllItem.Enabled = True
End Sub

Private Sub cmdPost_Click(Index As Integer)
'If IsNull(Me!cboPeriod) Then
'    MsgBox ("Enter a quarter to post this transaction to.")
'    Exit Sub
'End If


'Me!txtPostDate = Me!StartDate 'Make sure new record for qtr balances have date
'Me.Refresh


'If Me!txtPostedYN = 0 Then
'    Me!txtPostedYN = -1
'    Me!txtCheckDate = Me!StartDate
'    Me!txtStartDate = Me!StartDate
'    Me!txtDate = Me!StartDate
'    Me!txtCheckType = "Beginning Balance"
'    Me!EmpID = Me!txtEmployee
'    Me!txtName = Me!txtEmployeeName
'    Me!txtChecknumber = 999999999
'End If

'txtGross = GrossOld
'txtFICA = FICAOld
'txtFIT = FITOld
'txtState = StateOld
'txtLocal = LocalOld
'txtRegHrs = RegHrsOld
'txtOTHrs = OTHrsOld

'GrossOld.BackColor = 12632256
'FICAOld.BackColor = 12632256
'FITOld.BackColor = 12632256
'StateOld.BackColor = 12632256
'LocalOld.BackColor = 12632256
'RegHrsOld.BackColor = 12632256
'OTHrsOld.BackColor = 12632256


'Me.Refresh

' DoCmd.SetWarnings False
' DoCmd.OpenQuery ("Pyrl - BegBalYTDUpdate")

 
' DoCmd.OpenQuery ("Pyrl - ItemBegBalAppend")
' DoCmd.SetWarnings True
    
'   Set db = CurrentDb
'db.Execute ("Delete * from [Pyrl - ItemsBegBalWork]")
'DoCmd.SetWarnings False
'DoCmd.OpenQuery ("Pyrl - ItemBegBalUnmatched")
'DoCmd.SetWarnings True
'Me.Requery
'Forms![Pyrl - BeginningBalances]![Pyrl - BeginningBalSubform2].Form.Requery
    
'lblPostedYN.Caption = "Yes"
'cmdPost.Caption = "Update"

End Sub

Private Sub cmdPyrlRefresh_Click()
  'This is only needed for multi user apps
On Error GoTo RefreshErr
    With ADOCreatePay
        If .EditMode = adEditInProgress Then .CancelUpdate
        .Requery
    End With
  Exit Sub
RefreshErr:
  MsgBox Err.Description

End Sub

Private Sub cmdPyrlUpdate_Click()
  On Error GoTo UpdateErr
  
  With ADOCreatePay
  .Update
  .Requery
  End With
  Exit Sub
UpdateErr:
  MsgBox Err.Description

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
  'On Error GoTo UpdateErr
  
  'With ADOprimaryrs
  '.Update
  '.Requery
  'End With
  Dim oTxt As TextBox
    For Each oTxt In Me.txtfields
        If oTxt.Text = "" And oTxt.DataField <> "" Then
            If ADOemployee("" & oTxt.DataField & "").Type = 203 Or ADOemployee("" & oTxt.DataField & "").Type = 202 Then oTxt.Text = " "
        End If
    Next
  'MsgBox txtFields(1).Text
  ADOemployee![EMP ID] = txtfields(1).Text
  ADOemployee![EMP ID 2] = txtfields(1).Text
  ADOemployee.Update
  
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmdBeginning_Click()
    If txtfields(1) = "" Then
        MsgBox "Select an employee first.", vbInformation, "Error"
    Else
        txt(1) = ""
        txt(2) = ""
        picButtons.Visible = False
        txt(0) = txtfields(0)
        txtNew(0) = txtfields(1)
        Dim i As Integer
            For i = 1 To txtNew.UBound
                txtNew(i) = 0
            Next
        If txtBeginning(8) = "Posted" Then
            For i = 1 To txtOld.UBound
                txtOld(i) = txtBeginning(1)
            Next
        Else
            For i = 1 To txtOld.UBound
                txtOld(i) = 0
            Next
        End If
        frEmployeeSetup.Visible = False
        frPayrollItems.Visible = False
        frBeginningBal.ZOrder 0
        frBeginningBal.Visible = True
        Me.Caption = "Employee Setup - Beginning Balances"
        Form_Resize
    End If
End Sub

Private Sub Form_Load()
On Error GoTo FormErr
ShowStatus True
  Set db = New ADODB.Connection
  db.CursorLocation = adUseClient
  db.Open gblADOProvider

  'Select [ApplyItem],[Description],[ItemAmount],[ItemPercent],[Basis],[WageLow],[WageHigh],[YTDMax],[Account] from [Pyrl - Select Pyrl Items Work] where Type = 'Addition' order by [PyrlItemId]
  Call ResetPyrlItems
  grdDataSource "Select [ApplyItem],[Description],[ItemAmount],[ItemPercent],[Basis],[WageLow],[WageHigh],[YTDMax],[Account] from [Pyrl - Select Pyrl Items Work] where Type = 'Addition' order by [PyrlItemId]"

  Set grdDataGrid.DataSource = ADOprimaryrs
  picOptions(0).ZOrder 0
  mbDataChanged = False
  
  GetTextColor Me
  ShowStatus False
Exit Sub
FormErr:
  MsgBox Err.Description
  ShowStatus False
End Sub

Private Sub grdDataSource(SQLstatement As String)
  
  If ADOprimaryrs Is Nothing Then
  Else
    ADOprimaryrs.Close
  End If
  Set grdDataGrid.DataSource = Nothing
  Set ADOprimaryrs = New ADODB.Recordset
  ADOprimaryrs.Open SQLstatement, db, adOpenKeyset, adLockOptimistic, adCmdText

  Set grdDataGrid.DataSource = ADOprimaryrs

End Sub


Public Sub ResetPyrlItems()

'Clear Work Table
db.Execute "Delete * from [Pyrl - Select Pyrl Items Work]"
db.Execute "INSERT INTO [Pyrl - Select Pyrl Items Work] Select * from [Pyrl - Payroll Items]"

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If mbAddNewFlag Then Exit Sub

  Select Case KeyCode
    Case vbKeyEscape
      Unload Me
  End Select
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
  
If frPayrollItems.Visible = True Then
    frPrimary.Width = 9735
    frPrimary.Height = 6495
    grdDataGrid.Visible = False
    Me.Width = frPrimary.Width + (11310 - 11175)
    Me.Height = frPrimary.Height + (8475 - 7695) + 460
ElseIf frEmployeeSetup.Visible = True Then
    frPrimary.Width = 11175
    frPrimary.Height = 7455
    grdDataGrid.Visible = True
    Me.Width = frPrimary.Width + (11310 - 11175)
    Me.Height = frPrimary.Height + (8475 - 7695) + 460
ElseIf frBeginningBal.Visible = True Then
    frPrimary.Width = 11175
    frPrimary.Height = 7695
    grdDataGrid.Visible = True
    Me.Width = frPrimary.Width + (11310 - 11175)
    Me.Height = frPrimary.Height + (8475 - 7695) + 120
End If
        
SkipResize:
    If frPayrollItems.Visible = True Then
        frPrimary.Width = 9735
        frPrimary.Height = 6495
        grdDataGrid.Visible = False
    ElseIf frEmployeeSetup.Visible = True Then
        frPrimary.Width = 11175
        frPrimary.Height = 7455
        grdDataGrid.Visible = True
    ElseIf frBeginningBal.Visible = True Then
        frPrimary.Width = 11175
        frPrimary.Height = 7695
        grdDataGrid.Visible = True
    End If

  frPrimary.Left = (Me.ScaleWidth - frPrimary.Width) / 2
  lblTop.Caption = Me.Caption
  lblTop.Left = frPrimary.Left
  lblTop.Width = frPrimary.Width
  If frBeginningBal.Visible = True Then
    frPrimary.Top = (Me.ScaleHeight - frPrimary.Height) / 2 + 230
  Else
    frPrimary.Top = (Me.ScaleHeight - frPrimary.Height - picButtons.Height) / 2 + 230
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo FormErr
ShowStatus True
Dim rs As ADODB.Recordset

Set rs = New ADODB.Recordset
rs.Open "Select * FROM [Pyrl - Employees]", db, adOpenKeyset, adLockOptimistic, adCmdText

EndLoad db, rs, "Employees"

If txtfields(1).Text <> "" Then
  'updates the checklist Employees
      If ADOemployee.RecordCount > 0 Then
        If ADOemployee.EditMode <> 0 Then
          ADOemployee.CancelUpdate
        End If
      End If
      ADOemployee.Close
      Set ADOemployee = Nothing
End If
      rs.Close
      Set rs = Nothing
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

Private Sub grdDataGrid_AfterColEdit(ByVal ColIndex As Integer)
    If grdDataGrid.Row = -1 Or grdDataGrid.Columns(0) = "" Then Exit Sub
      SendKeys ("{ENTER}")
  If grdDataGrid.Row > 0 Then
      SendKeys ("{up}")
      SendKeys ("{down}")
  ElseIf grdDataGrid.Row = 0 Then
      SendKeys ("{down}")
      SendKeys ("{up}")
  End If
End Sub

Private Sub grdDataGrid_ButtonClick(ByVal ColIndex As Integer)
On Error GoTo Error_ButtClick
If grdDataGrid.Row < 0 Then Exit Sub
Select Case ColIndex
Case 0
         SendKeys ("{ENTER}")
   If grdDataGrid.Columns(0).Text = "No" Then
      grdDataGrid.Columns(0).Text = "Yes"
   Else
      grdDataGrid.Columns(0).Text = "No"
   End If
         SendKeys ("{ENTER}")
         SendKeys ("{down}")
         SendKeys ("{up}")
End Select
grdDataGrid_AfterColEdit 0
Exit Sub
Error_ButtClick:
    MsgBox "Please click the Table box before clicking the button"
End Sub

Private Sub grdDataGrid_Error(ByVal DataError As Integer, Response As Integer)
    If DataGridKnownError(DataError) Then
        Response = 0
    End If
End Sub

Private Sub grdDataGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
     grdDataGrid.col = 0
End Sub


Private Sub optAE_Click(Index As Integer)
Dim Selection As Integer
Select Case Index
Case 0
    Selection = 1
Case 1
    Selection = 2
End Select
If ADOemployee Is Nothing Then
Else
    ADOemployee![EMP Method] = Selection
End If
End Sub

Private Sub tbEmployee_Click()
    Dim i As Integer
    'show and enable the selected tab's controls
    'and hide and disable all others
    For i = 0 To tbEmployee.Tabs.count - 1
        If i = tbEmployee.SelectedItem.Index - 1 Then
            picOptions(i).Left = 4080
            picOptions(i).Top = 600
            picOptions(i).Enabled = True
            picOptions(i).ZOrder 0
            If i = 0 Then
                ' "Select * from [Pyrl - Select Pyrl Items Work] where Type = 'Addition' order by [PyrlItemId]"
                grdDataSource "Select [ApplyItem],[Description],[ItemAmount],[ItemPercent],[Basis],[WageLow],[WageHigh],[YTDMax],[Account] from [Pyrl - Select Pyrl Items Work] where Type = 'Addition' order by [PyrlItemId]"
                grdDataGrid.Columns(8).Caption = "Debit Account"
            ElseIf i = 1 Then
                ' = "SELECT * FROM [Pyrl - Select Pyrl Items Work]WHERE ((([Pyrl - Select Pyrl Items Work].Type) Like 'State Tax' Or ([Pyrl - Select Pyrl Items Work].Type) Like 'Local Tax' Or ([Pyrl - Select Pyrl Items Work].Type) Like 'Deduction'))ORDER BY [Pyrl - Select Pyrl Items Work].PyrlItemID"
                grdDataSource "Select [ApplyItem],[Description],[ItemAmount],[ItemPercent],[Basis],[WageLow],[WageHigh],[YTDMax],[Account] from [Pyrl - Select Pyrl Items Work] WHERE ((([Pyrl - Select Pyrl Items Work].Type) Like 'State Tax' Or ([Pyrl - Select Pyrl Items Work].Type) Like 'Local Tax' Or ([Pyrl - Select Pyrl Items Work].Type) Like 'Deduction'))ORDER BY [Pyrl - Select Pyrl Items Work].PyrlItemID"
                grdDataGrid.Columns(8).Caption = "Credit Account"
            End If
        Else
            picOptions(i).Left = -20000
            picOptions(i).Enabled = False
        End If
    Next
End Sub

Private Sub TypeDesc()
Dim txtmessage As String

Select Case cbPyrllItems(0)
Case "Addition"
    lblPyrllItems(20).Caption = "Debit Account"
   txtmessage = ""
   Select Case cbPyrllItems(1)
        Case "Gross"
            txtmessage = "Taxable Income"
        Case "AGI"
            txtmessage = "NonTaxable Income"
        Case "Net"
            txtmessage = "Reimbursement"
    End Select

Case "Deduction"
    lblPyrllItems(20).Caption = "Credit Account"
     txtmessage = ""
    Select Case cbPyrllItems(1)
        Case "Gross"
            txtmessage = "NonTaxable Deduction"
        Case "AGI"
            txtmessage = "AfterTax Deduction"
        Case "Net"
            txtmessage = "AfterTax Deduction"
     End Select

Case "State Tax"
    lblPyrllItems(20).Caption = "Credit Account"
    txtmessage = "State Tax"
    

Case "Local Tax"
    lblPyrllItems(20).Caption = "Credit Account"
    txtmessage = "Local Tax"
    

Case ""
    lblPyrllItems(20).Caption = "Account"
    txtmessage = ""
End Select
    txtPyrllItems(5) = txtmessage
End Sub

Private Sub txtFields_LostFocus(Index As Integer)
If Index = 1 Then
    OpenDBemp
End If
End Sub

Private Sub txtOld_KeyPress(Index As Integer, KeyAscii As Integer)
    keyResponse = CtrlValidate(KeyAscii, "")
    If keyResponse = True Then
    Else
       KeyAscii = 0
    End If
End Sub

Private Sub txtOld_LostFocus(Index As Integer)
    txtNew(Index) = txtOld(Index) - NZ(txtBeginning(Index), 0)
End Sub

Public Sub CallByUserEmpID(EmpID As String)
    Me.Show
    txtfields(1) = EmpID
    OpenDBemp
End Sub

