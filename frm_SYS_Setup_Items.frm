VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_SYS_Setup_Items 
   Caption         =   "Setup Items"
   ClientHeight    =   7530
   ClientLeft      =   2760
   ClientTop       =   3030
   ClientWidth     =   10380
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7530
   ScaleWidth      =   10380
   Begin VB.Frame frPrimary 
      Height          =   6375
      Left            =   0
      TabIndex        =   26
      Top             =   480
      Width           =   10335
      Begin VB.CommandButton cmdPricing 
         Caption         =   "Pricing"
         Height          =   735
         Index           =   0
         Left            =   9120
         Picture         =   "frm_SYS_Setup_Items.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   5400
         Width           =   855
      End
      Begin VB.CommandButton cmdTrx 
         Caption         =   "View Trx"
         Height          =   735
         Index           =   1
         Left            =   8040
         Picture         =   "frm_SYS_Setup_Items.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   5400
         Width           =   855
      End
      Begin VB.Frame Frame1 
         Height          =   1695
         Left            =   120
         TabIndex        =   49
         Top             =   4560
         Width           =   5295
         Begin VB.TextBox txtPrice 
            Alignment       =   2  'Center
            DataField       =   "INV BREAK Unit"
            DataSource      =   "adoPrimaryRS"
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   53
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox txtPrice 
            Alignment       =   2  'Center
            DataField       =   "INV BREAK Qty"
            DataSource      =   "adoPrimaryRS"
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   52
            Top             =   1200
            Width           =   1215
         End
         Begin VB.TextBox txtPrice 
            DataField       =   "INV BREAK Price"
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
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   3480
            Locked          =   -1  'True
            TabIndex        =   51
            Top             =   1200
            Width           =   1215
         End
         Begin VB.TextBox txtPrice 
            DataField       =   "INV BREAK Amount"
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
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   3480
            Locked          =   -1  'True
            TabIndex        =   50
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Unit"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   4
            Left            =   1080
            TabIndex        =   57
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Qty"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   5
            Left            =   1080
            TabIndex        =   56
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Amount Each"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   6
            Left            =   3480
            TabIndex        =   55
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Selling Price"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   19
            Left            =   3480
            TabIndex        =   54
            Top             =   960
            Width           =   1215
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1335
         Left            =   120
         TabIndex        =   44
         Top             =   2760
         Width           =   5295
         Begin VB.CommandButton btItemCostOfSales 
            Height          =   285
            Left            =   4080
            Picture         =   "frm_SYS_Setup_Items.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   960
            Width           =   375
         End
         Begin VB.CommandButton btItemInv 
            Height          =   285
            Left            =   4080
            Picture         =   "frm_SYS_Setup_Items.frx":091E
            Style           =   1  'Graphical
            TabIndex        =   45
            Top             =   600
            Width           =   375
         End
         Begin VB.CommandButton btItemSales 
            Height          =   285
            Left            =   4080
            Picture         =   "frm_SYS_Setup_Items.frx":0C28
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txtFields 
            DataField       =   "INV ITEM Sales Account"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   8
            Left            =   2400
            TabIndex        =   1
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox txtFields 
            DataField       =   "INV ITEM Inventory Account"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   9
            Left            =   2400
            TabIndex        =   3
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox txtFields 
            DataField       =   "INV ITEM Cost of Sales Account"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   10
            Left            =   2400
            TabIndex        =   4
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label lblfields 
            Caption         =   "GL Sales Account"
            Height          =   255
            Index           =   8
            Left            =   360
            TabIndex        =   48
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label lblfields 
            Caption         =   "GL Inventory Account"
            Height          =   255
            Index           =   9
            Left            =   360
            TabIndex        =   47
            Top             =   600
            Width           =   2055
         End
         Begin VB.Label lblfields 
            Caption         =   "GL Cost Of Sales Account"
            Height          =   255
            Index           =   10
            Left            =   360
            TabIndex        =   46
            Top             =   960
            Width           =   2055
         End
      End
      Begin VB.Frame Frame4 
         Height          =   2175
         Left            =   5520
         TabIndex        =   43
         Top             =   120
         Width           =   4695
         Begin VB.PictureBox Picture2 
            BorderStyle     =   0  'None
            Height          =   1935
            Left            =   120
            ScaleHeight     =   1935
            ScaleWidth      =   4455
            TabIndex        =   92
            Top             =   120
            Width           =   4455
            Begin VB.TextBox txtFields 
               DataField       =   "INV ITEM Reorder Qty"
               DataSource      =   "adoPrimaryRS"
               Enabled         =   0   'False
               Height          =   285
               Index           =   14
               Left            =   1560
               Locked          =   -1  'True
               TabIndex        =   96
               Top             =   1080
               Width           =   1215
            End
            Begin VB.TextBox txtFields 
               DataField       =   "INV ITEM Vendors Number"
               DataSource      =   "adoPrimaryRS"
               Height          =   285
               Index           =   7
               Left            =   1560
               TabIndex        =   95
               Top             =   720
               Width           =   2295
            End
            Begin VB.TextBox txtFields 
               DataField       =   "INV ITEM Vendor ID"
               DataSource      =   "adoPrimaryRS"
               Height          =   285
               Index           =   6
               Left            =   1560
               Locked          =   -1  'True
               TabIndex        =   94
               Top             =   360
               Width           =   1935
            End
            Begin VB.CommandButton btItemVen 
               Height          =   285
               Left            =   3480
               Picture         =   "frm_SYS_Setup_Items.frx":0F32
               Style           =   1  'Graphical
               TabIndex        =   93
               Top             =   360
               Width           =   375
            End
            Begin VB.Label lblLabels 
               Alignment       =   1  'Right Justify
               Caption         =   "Reorder Qty"
               DataSource      =   "adoPrimaryRS"
               Height          =   255
               Index           =   13
               Left            =   240
               TabIndex        =   99
               Top             =   1080
               Width           =   1215
            End
            Begin VB.Label lblLabels 
               Alignment       =   1  'Right Justify
               Caption         =   "Vendor Number"
               Height          =   255
               Index           =   12
               Left            =   240
               TabIndex        =   98
               Top             =   720
               Width           =   1215
            End
            Begin VB.Label lblLabels 
               Alignment       =   1  'Right Justify
               Caption         =   "Vendor ID"
               Height          =   255
               Index           =   11
               Left            =   240
               TabIndex        =   97
               Top             =   360
               Width           =   1215
            End
         End
         Begin VB.PictureBox Picture3 
            BorderStyle     =   0  'None
            Height          =   1815
            Left            =   120
            ScaleHeight     =   1815
            ScaleWidth      =   4455
            TabIndex        =   100
            Top             =   240
            Width           =   4455
            Begin MSDataGridLib.DataGrid DataGrid1 
               Bindings        =   "frm_SYS_Setup_Items.frx":123C
               Height          =   1815
               Left            =   120
               TabIndex        =   101
               Top             =   0
               Width           =   4215
               _ExtentX        =   7435
               _ExtentY        =   3201
               _Version        =   393216
               AllowUpdate     =   -1  'True
               HeadLines       =   1
               RowHeight       =   15
               FormatLocked    =   -1  'True
               AllowAddNew     =   -1  'True
               AllowDelete     =   -1  'True
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
               Caption         =   "Inventory Production Items"
               ColumnCount     =   6
               BeginProperty Column00 
                  DataField       =   "INV KIT ID"
                  Caption         =   "INV KIT ID"
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
                  DataField       =   "INV KIT Item ID"
                  Caption         =   "INV KIT Item ID"
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
                  DataField       =   "INV KIT Sub Item ID"
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
               BeginProperty Column03 
                  DataField       =   "INV KIT Qty"
                  Caption         =   "Quantity"
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
                  DataField       =   "INV KIT Unit"
                  Caption         =   "Unit"
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
                  DataField       =   "INV KIT Cost"
                  Caption         =   "INV KIT Cost"
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
                     Object.Visible         =   0   'False
                     ColumnWidth     =   915.024
                  EndProperty
                  BeginProperty Column01 
                     Object.Visible         =   0   'False
                     ColumnWidth     =   1739.906
                  EndProperty
                  BeginProperty Column02 
                     Button          =   -1  'True
                     ColumnWidth     =   1739.906
                  EndProperty
                  BeginProperty Column03 
                     ColumnWidth     =   840.189
                  EndProperty
                  BeginProperty Column04 
                     ColumnWidth     =   945.071
                  EndProperty
                  BeginProperty Column05 
                     Object.Visible         =   0   'False
                     ColumnWidth     =   1739.906
                  EndProperty
               EndProperty
            End
         End
      End
      Begin VB.Frame Frame5 
         Height          =   2415
         Left            =   5520
         TabIndex        =   36
         Top             =   2760
         Width           =   4695
         Begin VB.TextBox txtFields 
            DataField       =   "INV ITEM Costing Method"
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
            Index           =   15
            Left            =   1440
            TabIndex        =   6
            Top             =   480
            Width           =   735
         End
         Begin VB.TextBox txtFields 
            DataField       =   "INV ITEM Standard Cost"
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
            Index           =   16
            Left            =   1440
            TabIndex        =   7
            Top             =   840
            Width           =   1455
         End
         Begin VB.TextBox txtFields 
            DataField       =   "INV ITEM Last Cost"
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
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   1440
            Width           =   1095
         End
         Begin VB.TextBox txtFields 
            DataField       =   "INV ITEM Average Cost"
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
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   1800
            Width           =   1095
         End
         Begin VB.TextBox txtFields 
            DataField       =   "INV ITEM Qty On Hand"
            DataSource      =   "adoPrimaryRS"
            Enabled         =   0   'False
            Height          =   285
            Index           =   12
            Left            =   3480
            Locked          =   -1  'True
            TabIndex        =   10
            Top             =   1440
            Width           =   1095
         End
         Begin VB.TextBox txtFields 
            DataField       =   "INV ITEM Qty On Order"
            DataSource      =   "adoPrimaryRS"
            Enabled         =   0   'False
            Height          =   285
            Index           =   13
            Left            =   3480
            Locked          =   -1  'True
            TabIndex        =   11
            Top             =   1800
            Width           =   1095
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Costing Method"
            DataSource      =   "adoPrimaryRS"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   42
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Standard Cost"
            DataSource      =   "adoPrimaryRS"
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   41
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Last Cost"
            DataSource      =   "adoPrimaryRS"
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   40
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Average Cost"
            DataSource      =   "adoPrimaryRS"
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   39
            Top             =   1800
            Width           =   975
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Qty On Hand"
            DataSource      =   "adoPrimaryRS"
            Height          =   255
            Index           =   14
            Left            =   2280
            TabIndex        =   38
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Qty On Order"
            DataSource      =   "adoPrimaryRS"
            Height          =   255
            Index           =   15
            Left            =   2280
            TabIndex        =   37
            Top             =   1800
            Width           =   1095
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Price List "
         Height          =   735
         Left            =   6960
         Picture         =   "frm_SYS_Setup_Items.frx":1251
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   5400
         Width           =   855
      End
      Begin VB.Frame frPriceList 
         Height          =   6135
         Left            =   120
         TabIndex        =   27
         Top             =   120
         Visible         =   0   'False
         Width           =   10095
         Begin VB.CommandButton cmdAddNew 
            Caption         =   "New Item"
            Height          =   735
            Left            =   7080
            Picture         =   "frm_SYS_Setup_Items.frx":155B
            Style           =   1  'Graphical
            TabIndex        =   91
            Top             =   600
            Width           =   855
         End
         Begin VB.CheckBox optPriceList 
            Caption         =   "Cost Price"
            Height          =   255
            Index           =   3
            Left            =   480
            TabIndex        =   34
            Top             =   720
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox optPriceList 
            Caption         =   "Vendor ID"
            Height          =   255
            Index           =   2
            Left            =   480
            TabIndex        =   33
            Top             =   360
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox optPriceList 
            Caption         =   "Quantity On Hand"
            Height          =   255
            Index           =   5
            Left            =   2880
            TabIndex        =   32
            Top             =   360
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.CheckBox optPriceList 
            Caption         =   "Quantity On Order"
            Height          =   255
            Index           =   6
            Left            =   2880
            TabIndex        =   31
            Top             =   720
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.CommandButton cmdBackPriceList 
            Caption         =   "Back"
            Height          =   735
            Left            =   9120
            Picture         =   "frm_SYS_Setup_Items.frx":1865
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   600
            Width           =   855
         End
         Begin VB.CommandButton cmdPrint 
            Caption         =   "Print"
            Height          =   735
            Left            =   8040
            Picture         =   "frm_SYS_Setup_Items.frx":1B6F
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   600
            Width           =   855
         End
         Begin VB.CheckBox optPriceList 
            Caption         =   "Quantity On Order"
            Height          =   255
            Index           =   7
            Left            =   4920
            TabIndex        =   28
            Top             =   360
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin MSDataGridLib.DataGrid grdDatagrid3 
            Bindings        =   "frm_SYS_Setup_Items.frx":1E79
            Height          =   4455
            Left            =   120
            TabIndex        =   35
            Top             =   1560
            Width           =   9855
            _ExtentX        =   17383
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
            Caption         =   "Price List"
            ColumnCount     =   8
            BeginProperty Column00 
               DataField       =   "INV ITEM Id"
               Caption         =   "Id"
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
               DataField       =   "INV ITEM Description"
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
               DataField       =   "INV ITEM Vendor ID"
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
            BeginProperty Column03 
               DataField       =   "INV ITEM Last Cost"
               Caption         =   "Cost"
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
               DataField       =   "INV ITEM Price"
               Caption         =   "Selling Price"
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
               DataField       =   "INV ITEM Qty On Hand"
               Caption         =   "On Hand"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column06 
               DataField       =   "INV ITEM Qty On Order"
               Caption         =   "On Order"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column07 
               DataField       =   "INV ITEM Notes"
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
                  ColumnWidth     =   1035.213
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   2039.811
               EndProperty
               BeginProperty Column02 
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   1049.953
               EndProperty
               BeginProperty Column03 
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   945.071
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   945.071
               EndProperty
               BeginProperty Column05 
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   959.811
               EndProperty
               BeginProperty Column06 
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   959.811
               EndProperty
               BeginProperty Column07 
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   1244.976
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame frPricing 
         Height          =   6135
         Left            =   120
         TabIndex        =   64
         Top             =   120
         Visible         =   0   'False
         Width           =   10095
         Begin VB.TextBox txtPrice 
            DataField       =   "INV BREAK Amount"
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
            Index           =   4
            Left            =   5880
            TabIndex        =   73
            Top             =   1080
            Width           =   1335
         End
         Begin VB.TextBox txtPrice 
            DataField       =   "INV BREAK Price"
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
            Index           =   5
            Left            =   5880
            Locked          =   -1  'True
            TabIndex        =   72
            Top             =   1440
            Width           =   1335
         End
         Begin VB.TextBox txtPrice 
            Alignment       =   2  'Center
            DataField       =   "INV BREAK Qty"
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
            Index           =   6
            Left            =   1800
            TabIndex        =   71
            Top             =   1080
            Width           =   975
         End
         Begin VB.TextBox txtPrice 
            Alignment       =   2  'Center
            DataField       =   "INV BREAK Unit"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   7
            Left            =   1800
            TabIndex        =   70
            Top             =   1440
            Width           =   975
         End
         Begin VB.TextBox txtPrice 
            DataField       =   "INV BREAK Description"
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
            Enabled         =   0   'False
            Height          =   285
            Index           =   8
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   69
            Top             =   720
            Width           =   2655
         End
         Begin VB.TextBox txtPrice 
            DataField       =   "INV BREAK ID"
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
            Enabled         =   0   'False
            Height          =   285
            Index           =   9
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   68
            Top             =   360
            Width           =   1455
         End
         Begin VB.ComboBox cbPrice 
            DataField       =   "INV BREAK Default"
            BeginProperty DataFormat 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "Yes"
               FalseValue      =   "No"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   7
            EndProperty
            Height          =   315
            ItemData        =   "frm_SYS_Setup_Items.frx":1E8E
            Left            =   5880
            List            =   "frm_SYS_Setup_Items.frx":1E98
            TabIndex        =   66
            Text            =   "No"
            Top             =   720
            Width           =   1335
         End
         Begin MSDataGridLib.DataGrid grdDatagrid1 
            Bindings        =   "frm_SYS_Setup_Items.frx":1EA5
            Height          =   4095
            Left            =   120
            TabIndex        =   74
            Top             =   1920
            Width           =   9855
            _ExtentX        =   17383
            _ExtentY        =   7223
            _Version        =   393216
            AllowUpdate     =   0   'False
            HeadLines       =   1
            RowHeight       =   15
            FormatLocked    =   -1  'True
            AllowDelete     =   -1  'True
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
            Caption         =   "Item Pricing"
            ColumnCount     =   7
            BeginProperty Column00 
               DataField       =   "INV BREAK Default"
               Caption         =   "Default"
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
               DataField       =   "INV BREAK ID"
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
            BeginProperty Column02 
               DataField       =   "INV BREAK Description"
               Caption         =   "Item Description"
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
               DataField       =   "INV BREAK Qty"
               Caption         =   "Qty"
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
               DataField       =   "INV BREAK Unit"
               Caption         =   "Unit"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column05 
               DataField       =   "INV BREAK Amount"
               Caption         =   "Amount Unit"
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
               DataField       =   "INV BREAK Price"
               Caption         =   "Selling Price"
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
                  ColumnWidth     =   689.953
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1335.118
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   2775.118
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   840.189
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   945.071
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   1335.118
               EndProperty
               BeginProperty Column06 
                  ColumnWidth     =   1335.118
               EndProperty
            EndProperty
         End
         Begin VB.CommandButton cmdBackPricing 
            Caption         =   "&Back"
            Height          =   735
            Left            =   9120
            Picture         =   "frm_SYS_Setup_Items.frx":1EB5
            Style           =   1  'Graphical
            TabIndex        =   65
            Top             =   1080
            Width           =   855
         End
         Begin VB.CommandButton cmdAddPricing 
            Caption         =   "&New Price"
            Height          =   735
            Left            =   9120
            Picture         =   "frm_SYS_Setup_Items.frx":21BF
            Style           =   1  'Graphical
            TabIndex        =   67
            Top             =   240
            Width           =   855
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Selling Price:"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   20
            Left            =   4320
            TabIndex        =   81
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Amount Unit:"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   21
            Left            =   4440
            TabIndex        =   80
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Qty:"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   22
            Left            =   840
            TabIndex        =   79
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Unit:"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   23
            Left            =   720
            TabIndex        =   78
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Item Description:"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   24
            Left            =   240
            TabIndex        =   77
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Item ID:"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   25
            Left            =   240
            TabIndex        =   76
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Default:"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   26
            Left            =   4800
            TabIndex        =   75
            Top             =   720
            Width           =   975
         End
      End
      Begin VB.Frame frTranx 
         Height          =   6135
         Left            =   120
         TabIndex        =   58
         Top             =   120
         Visible         =   0   'False
         Width           =   10095
         Begin VB.CommandButton cmdTranx 
            Caption         =   "Back"
            Height          =   735
            Left            =   9120
            Picture         =   "frm_SYS_Setup_Items.frx":24C9
            Style           =   1  'Graphical
            TabIndex        =   63
            Top             =   240
            Width           =   855
         End
         Begin VB.CheckBox chkTranx 
            Caption         =   "Sales"
            Height          =   255
            Index           =   0
            Left            =   2040
            TabIndex        =   61
            Top             =   360
            Width           =   1575
         End
         Begin VB.CheckBox chkTranx 
            Caption         =   "Purchases"
            Height          =   255
            Index           =   1
            Left            =   4080
            TabIndex        =   60
            Top             =   360
            Width           =   1575
         End
         Begin VB.CheckBox chkTranx 
            Caption         =   "Inventory"
            Height          =   255
            Index           =   2
            Left            =   6480
            TabIndex        =   59
            Top             =   360
            Width           =   1575
         End
         Begin MSDataGridLib.DataGrid grdDatagrid2 
            Bindings        =   "frm_SYS_Setup_Items.frx":27D3
            Height          =   4935
            Left            =   120
            TabIndex        =   62
            Top             =   1080
            Width           =   9855
            _ExtentX        =   17383
            _ExtentY        =   8705
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
            ColumnCount     =   5
            BeginProperty Column00 
               DataField       =   "date"
               Caption         =   "date"
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
            BeginProperty Column01 
               DataField       =   "type"
               Caption         =   "type"
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
               DataField       =   "qty"
               Caption         =   "qty"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column03 
               DataField       =   "cost"
               Caption         =   "cost"
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
               DataField       =   "Doc #"
               Caption         =   "Doc #"
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
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   2039.811
               EndProperty
               BeginProperty Column02 
                  Alignment       =   2
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column03 
                  Alignment       =   1
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   2039.811
               EndProperty
            EndProperty
         End
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "GL Accounts"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   84
         Top             =   2400
         Width           =   5295
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Item Default Price"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   83
         Top             =   4200
         Width           =   5295
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   4
         Left            =   5520
         TabIndex        =   82
         Top             =   2400
         Width           =   4695
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   10335
      TabIndex        =   85
      Top             =   0
      Width           =   10335
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vendor"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   1
         Left            =   5520
         TabIndex        =   88
         Top             =   120
         Width           =   4695
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Item"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   86
         Top             =   120
         Width           =   5295
      End
      Begin VB.Label lblTranx 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Item Transaction"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   90
         Top             =   120
         Visible         =   0   'False
         Width           =   10335
      End
      Begin VB.Label lblPriceList 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Price List"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   89
         Top             =   120
         Visible         =   0   'False
         Width           =   10335
      End
      Begin VB.Label lblPricing 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Item Pricing"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   87
         Top             =   120
         Visible         =   0   'False
         Width           =   10335
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
      ScaleWidth      =   10380
      TabIndex        =   25
      Top             =   6930
      Width           =   10380
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   300
         Left            =   4320
         TabIndex        =   19
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   300
         Left            =   3240
         TabIndex        =   18
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   300
         Left            =   2160
         TabIndex        =   17
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   300
         Left            =   1080
         TabIndex        =   16
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   300
         Left            =   0
         TabIndex        =   15
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
      ScaleWidth      =   10380
      TabIndex        =   0
      Top             =   7230
      Width           =   10380
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Picture         =   "frm_SYS_Setup_Items.frx":27E3
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Picture         =   "frm_SYS_Setup_Items.frx":2B25
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Picture         =   "frm_SYS_Setup_Items.frx":2E67
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Picture         =   "frm_SYS_Setup_Items.frx":31A9
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   24
         Top             =   0
         Width           =   3360
      End
   End
End
Attribute VB_Name = "frm_SYS_Setup_Items"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ADOprimaryrs As ADODB.Recordset
Attribute ADOprimaryrs.VB_VarHelpID = -1
Dim ADOsecondaryRS As ADODB.Recordset
Dim ADOThirdRS As ADODB.Recordset
Dim ADOfourthRS As ADODB.Recordset
Dim ADOItemRS As ADODB.Recordset
Dim db As ADODB.Connection

Dim SecondaryLoad As Boolean
Dim ThirdLoad As Boolean

Dim CurrItem As String

Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean

Private Sub btcbRefresh_Click(Index As Integer)
    Dim tmp As String
    tmp = cbfields(Index).Text
    loadCombo Index
    cbfields(Index).Text = tmp
End Sub

Private Sub btItemCostOfSales_Click()
    Dim No As Integer
    Dim sql As String
    Dim ghead As String
    Dim fhead As String
    
    No = 8
    sql = "select [GL COA Account No], [GL COA Account Name]" & _
        "from [GL Chart of Accounts]"
    ghead = "Account Description"
    fhead = "Account No//Description"
    
    AllLookup.GetWhichTable No, sql, ghead, fhead, db
    'AllLookup.Show vbModal
    txtFields(10).SetFocus
End Sub

Private Sub btItemID_Click()
    Dim ghead As String
    Dim fhead As String

    ghead = "Items"
    fhead = "ID//Description"
    AllLookup.ToWhichRecord ADOprimaryrs, ghead, fhead
    'AllLookup.Show vbModal
    ItemPrice "' AND [INV BREAK Default]=TRUE"
End Sub

Private Sub ItemPrice(Optional AddClause As String)
Dim oText As TextBox
    
    If ADOprimaryrs.EOF Or ADOprimaryrs.BOF Or mbAddNewFlag = True Then
        For Each oText In Me.txtPrice
            Set oText.DataSource = Nothing
            oText.Text = ""
        Next
            Set cbPrice.DataSource = Nothing
            cbPrice.Text = ""
    Else
        If SecondaryLoad = True Then
            ADOsecondaryRS.CancelUpdate
            ADOsecondaryRS.Close
            Set ADOsecondaryRS = Nothing
        End If
        If AddClause = "" Then
            AddClause = "'"
        End If
        Set ADOsecondaryRS = New ADODB.Recordset
        ADOsecondaryRS.Open "SELECT * FROM [INV Items Break] WHERE [INV BREAK ID]='" & ADOprimaryrs![INV ITEM Id] & AddClause, db, adOpenKeyset, adLockOptimistic, adCmdText
        SecondaryLoad = True
        'MsgBox "SELECT * FROM [INV Items Break] WHERE [INV BREAK ID]='" & ADOprimaryrs![INV ITEM Id] & AddClause
        'Bind the text boxes to the data provider
        If ADOsecondaryRS.RecordCount > 0 Then
            For Each oText In Me.txtPrice
              Set oText.DataSource = ADOsecondaryRS
            Next
              Set cbPrice.DataSource = ADOsecondaryRS
        Else
            If AddClause = "'" Then
                For Each oText In Me.txtPrice
                  Set oText.DataSource = ADOsecondaryRS
                Next
                  Set cbPrice.DataSource = ADOsecondaryRS
                cmdAddPricing_Click
            Else
                MsgBox "Either this is new item or maybe the Item pricing is empty", vbCritical, "Information"
                For Each oText In Me.txtPrice
                  Set oText.DataSource = Nothing
                  oText.Text = ""
                Next
                  Set cbPrice.DataSource = Nothing
                  cbPrice.Text = ""
            End If
        End If
    End If
End Sub

Private Sub btItemInv_Click()
    Dim No As Integer
    Dim sql As String
    Dim ghead As String
    Dim fhead As String
    
    No = 7
    sql = "select [GL COA Account No], [GL COA Account Name]" & _
        "from [GL Chart of Accounts]"
    ghead = "Account Description"
    fhead = "Account No//Description"
    
    AllLookup.GetWhichTable No, sql, ghead, fhead, db
    'AllLookup.Show vbModal
    txtFields(9).SetFocus
End Sub

Private Sub btItemSales_Click()
    Dim No As Integer
    Dim sql As String
    Dim ghead As String
    Dim fhead As String
    
    No = 6
    sql = "select [GL COA Account No], [GL COA Account Name]" & _
        "from [GL Chart of Accounts]"
    ghead = "Account Description"
    fhead = "Account No//Description"
    
    AllLookup.GetWhichTable No, sql, ghead, fhead, db
    'AllLookup.Show vbModal
    txtFields(8).SetFocus
End Sub

Private Sub btItemVen_Click()
    Dim No As Integer
    Dim sql As String
    Dim ghead As String
    Dim fhead As String
    
    No = 5
    sql = "select [AP VEN ID], [AP VEN Name] from [AP Vendor]"
    ghead = "Vendors"
    fhead = "ID//Name"
    
    AllLookup.GetWhichTable No, sql, ghead, fhead, db
    'AllLookup.Show vbModal
    txtFields(6).SetFocus

End Sub

Private Sub cbFields_Click(Index As Integer)
    If Index = 1 Then
        If cbfields(Index).Text = "Assembly" Then
            Picture3.ZOrder 0
        Else
            Picture2.ZOrder 0
        End If
    End If
End Sub

Private Sub cbfields_LostFocus(Index As Integer)
Select Case Index
Case 1
   CheckCombo cbfields(Index), "[Type]", "[LIST Item Types]", db, True
Case 2
   CheckCombo cbfields(Index), "[Category]", "[LIST Item Categories]", db, True
End Select
End Sub


Private Sub cbPrice_KeyPress(KeyAscii As Integer)
    keyResponse = CtrlValidate(KeyAscii, "")
    If keyResponse = True Then
    Else
       KeyAscii = 0
    End If
End Sub

Private Sub cbPrice_LostFocus()
   If CheckCombo(cbPrice) Then
        MsgBox "There is no such selection", vbInformation, "Information"
   End If
   ADOsecondaryRS.Update
End Sub

Private Sub chkFields_Click(Index As Integer)
If Index = 0 Then
    If Trim(txtPrice(0)) = "" Then
        chkFields(0).Value = 1
        Exit Sub
    End If
End If
End Sub

Private Sub chkTranx_Click(Index As Integer)
ShowStatus True
Dim SQLstatement As String
Dim Source As String

    SQLstatement = ""
    If chkTranx(0).Value = 1 Then
        SQLstatement = "'Invoice','Sales Memo','Return','Credit Memo'"
    End If
    If chkTranx(1).Value = 1 Then
      If SQLstatement = "" Then
        SQLstatement = "'Receiving','PO','Voucher','Credit Memo','RMA'"
      Else
        SQLstatement = SQLstatement & ",'Receiving','PO','Voucher','Credit Memo','RMA'"
      End If
    End If
    If chkTranx(2).Value = 1 Then
      If SQLstatement = "" Then
        SQLstatement = "'Increase','Decrease','Production'"
      Else
        SQLstatement = SQLstatement & ",'Increase','Decrease','Production'"
      End If
    End If
    If SQLstatement = "" Then
        Set grdDataGrid2.DataSource = Nothing
        grdDataGrid2.Refresh
        ShowStatus False
        Exit Sub
    End If
    'Debug.Print SQLStatement
    Source = "SELECT [Date],[Type],[Qty],[Cost],[Doc #],[Memo] "
    Source = Source & "FROM [qryInventoryTransactions] "
    Source = Source & "WHERE [ID] = '" & txtFields(0) & "' AND [Type] in (" & SQLstatement & ")"
    
    Set grdDataGrid2.DataSource = Nothing
    If ThirdLoad = True Then
        ADOThirdRS.Close
        Set ADOThirdRS = Nothing
    End If
    Set ADOThirdRS = New ADODB.Recordset
    ADOThirdRS.Open Source, db, adOpenKeyset, adLockOptimistic, adCmdText
    
    Set grdDataGrid2.DataSource = ADOThirdRS
ShowStatus False
End Sub

Private Sub cmdAddNew_Click()
    frPriceList.Visible = False
    lblPriceList.Visible = False
    SetButtons True
    cmdAdd_Click
    cmdBackPriceList.Caption = "Back"
End Sub

Private Sub cmdBackPricing_Click()
On Error GoTo FormErr
ShowStatus True
'Set grdDatagrid1.DataSource = Nothing

'On Error GoTo FormErr
Dim i As Integer
Dim CountYes As Integer

    
    'check for the default value and must not be more than one
    CountYes = 0
    If Trim(ADOsecondaryRS![INV BREAK Unit]) = "" Then
        ADOsecondaryRS![INV BREAK Unit] = "Unit"
        ADOsecondaryRS.Update
    'ADOsecondaryRS.CancelUpdate
'    ADOsecondaryRS.Requery
        ADOsecondaryRS.MoveFirst
    End If
    
    txtPrice(5).Text = CInt(txtPrice(6).Text) * CCur(txtPrice(4).Text)
    'If CCur(txtPrice(5).Text) = 0 Or ADOsecondaryRS![INV BREAK Default] = False Then
    '    ShowStatus False
    '    MsgBox "There is no default price for the item"
    '    Exit Sub
    'End If
    
    For i = 1 To ADOsecondaryRS.RecordCount
        If ADOsecondaryRS![INV BREAK Default] = True Then CountYes = CountYes + 1
        ADOsecondaryRS.MoveNext
    Next
        
    ADOsecondaryRS.MoveFirst
    
    ADOsecondaryRS.Find "[INV BREAK Default]=True"
    If CountYes = 1 And ADOsecondaryRS.EOF = False Then
        If CCur(txtPrice(5)) = 0 Then
            MsgBox "Item Price cannot be zero", vbInformation, "Information"
            ShowStatus False
            Exit Sub
        End If
        ADOprimaryrs![INV ITEM Unit] = ADOsecondaryRS![INV BREAK Unit] & ""
        ADOprimaryrs![INV ITEM Price] = ADOsecondaryRS![INV BREAK Price]
        ADOprimaryrs.Update
        'MsgBox ADOsecondaryRS![INV BREAK Unit] & "  " & ADOsecondaryRS![INV BREAK Price]
    ElseIf CountYes > 1 Then
        MsgBox "There is more than one default prive. Only one default price is allow.", vbCritical, "Error"
        ShowStatus False
        Exit Sub
    Else
        MsgBox "There is no default price. Only one default price is allow", vbCritical, "Error"
        ShowStatus False
        Exit Sub
    End If
    ADOprimaryrs![INV ITEM Inactive YN] = False
    ADOprimaryrs.Update
    
    ItemPrice "' AND [INV BREAK Default]=TRUE"
    frPricing.Visible = False
    lblPricing.Visible = False
    SetButtons True
    'Me.Width = 10200
    'Me.Height = 7650
    Me.Caption = "Setup Items"
  
  ShowStatus False
Exit Sub
FormErr:
  MsgBox Err.Description
  ShowStatus False
End Sub

Private Sub cmdPricing_Click(Index As Integer)
    If mbAddNewFlag Then
        MsgBox "Can't process your request during adding new data", vbInformation, "Information"
        Exit Sub
    End If
    If txtFields(0) = "" Or txtFields(1) = "" Then
        MsgBox "Please select Item ID first before continue or item description is empty", vbInformation, "Information"
        Exit Sub
    End If
    
    ItemPrice
    frPricing.Visible = True
    lblPricing.Visible = True
    lblPricing.ZOrder 0
    frPricing.ZOrder 0
    SetButtons False
    'Me.Height = 5790
    Me.Caption = "Setup Items - Item Pricing"
    'Set grdDatagrid1.DataSource = Nothing
    txtPrice(9) = txtFields(0)
    txtPrice(8) = txtFields(1)
    'If ADOsecondaryRS.RecordCount = 0 Then
    Set grdDataGrid1.DataSource = ADOsecondaryRS
    If ADOsecondaryRS.EditMode <> adEditAdd Then
        txtPrice(9) = txtFields(0)
        txtPrice(7) = "Unit"
    End If
End Sub

Private Sub cmdAddPricing_Click()
If cmdAddPricing.Caption = "&New Price" Then
  With ADOsecondaryRS
    .AddNew
    txtPrice(9) = txtFields(0)
    txtPrice(8) = txtFields(1)
    ![INV BREAK ID] = txtFields(0) & ""
    ![INV BREAK Description] = txtFields(8) & ""
    ![INV BREAK Unit] = "Unit"
    cbPrice.Text = "No"
    txtPrice(4) = "$0.00"
    txtPrice(5) = "$0.00"
    txtPrice(6) = 0
    .Update
    ADOsecondaryRS.Requery
  End With
End If
End Sub

Private Sub cmdTranx_Click()
    ADOThirdRS.Close
    Set ADOThirdRS = Nothing
    ThirdLoad = False
    frTranx.Visible = False
    lblTranx.Visible = False
    SetButtons True
    'Me.Height = 7635
    'Me.Width = 10200
End Sub

Private Sub cmdTrx_Click(Index As Integer)
    If mbAddNewFlag Then
        MsgBox "Can't process your request during adding new data", vbInformation, "Information"
        Exit Sub
    End If

ShowStatus True
Dim Source As String
Dim SQLstatement As String
    
    If txtFields(0) = "" Then
        MsgBox "Please select Item ID first before continue", vbInformation, "Information"
        ShowStatus False
        Exit Sub
    End If

    SQLstatement = " 'Invoice','Sales Memo','Return','Credit Memo','Receiving','PO','Voucher','Credit Memo','RMA','Increase','Decrease','Production'"
    
    If ThirdLoad = False Then
    
        Source = "SELECT [Date],[Type],[Qty],[Cost],[Doc #],[Memo] "
        Source = Source & "FROM [qryInventoryTransactions] "
        Source = Source & "WHERE [ID] = '" & txtFields(0) & "' AND [Type] in (" & SQLstatement & ")"

        Set ADOThirdRS = New ADODB.Recordset
        ADOThirdRS.Open Source, db, adOpenKeyset, adLockOptimistic, adCmdText
        'MsgBox ADOThirdRS.RecordCount
        Set grdDataGrid2.DataSource = ADOThirdRS
        ThirdLoad = True
    End If
    
    chkTranx(0).Value = 1
    chkTranx(1).Value = 1
    chkTranx(2).Value = 1
    
    frTranx.Visible = True
    lblTranx.Visible = True
    lblTranx.ZOrder 0
    frTranx.ZOrder 0
    'Me.Height = 5790
    SetButtons False
    
ShowStatus False
End Sub

Public Sub PriceList()
    cmdBackPriceList.Caption = "Close"
    frm_SYS_Setup_Items.Show
    Command1_Click
End Sub


Private Sub Command1_Click()
    If mbAddNewFlag Then
        MsgBox "Can't process your request during adding new data", vbInformation, "Information"
        Exit Sub
    End If
    
    Set ADOfourthRS = New ADODB.Recordset
    ADOfourthRS.Open "select [INV ITEM Id],[INV ITEM Description],[INV ITEM Vendor ID],[INV ITEM Last Cost],[INV ITEM Price],[INV ITEM Qty On Hand],[INV ITEM Qty On Order],[INV ITEM Notes] from [INV Items] Order by [INV ITEM Id]", db, adOpenKeyset, adLockOptimistic, adCmdText
    
    Set grdDataGrid3.DataSource = ADOfourthRS
    
    'Me.Height = 6990
    frPriceList.Visible = True
    frPriceList.ZOrder 0
    'frPriceList.Top = 480
    'frPriceList.Left = 0
    SetButtons False
    lblPriceList.Visible = True
    lblPriceList.ZOrder 0
End Sub

Private Sub cmdBackPriceList_Click()
    Set grdDataGrid3.DataSource = Nothing
    ADOfourthRS.Close
    Set ADOfourthRS = Nothing
If cmdBackPriceList.Caption = "Back" Then
    frPriceList.Visible = False
    lblPriceList.Visible = False
    SetButtons True
    'Me.Height = 7695
Else
    Unload Me
End If
End Sub

Private Sub DataGrid1_AfterColEdit(ByVal ColIndex As Integer)
If mbAddNewFlag = True Then Exit Sub
  If DataGrid1.Row = -1 Or DataGrid1.Columns(0) = "" Then Exit Sub
      SendKeys ("{ENTER}")
  If DataGrid1.Row > 0 Then
      SendKeys ("{up}")
      SendKeys ("{down}")
  ElseIf DataGrid1.Row = 0 Then
      SendKeys ("{down}")
      SendKeys ("{up}")
  End If
End Sub

Private Sub DataGrid1_BeforeDelete(Cancel As Integer)
    Dim DeleteCration As Integer
    
    DeleteCration = MsgBox("Attempting to delete the data. " & vbCr & "Are you sure?", vbYesNo, "Deleting Confirmation")
    If DeleteCration = vbNo Then Cancel = 1
End Sub

Private Sub DataGrid1_ButtonClick(ByVal ColIndex As Integer)
On Error GoTo Error_ButtClick
If mbAddNewFlag = True Then Exit Sub
If DataGrid1.Columns(0) <> "" Then grdOnAddNew = False
Select Case ColIndex
Case 2   'select the item from the ITEM_ID
    INV_ITEM
Case Else
End Select
If grdOnAddNew = True And DataGrid1.Columns(2) <> "" Then
    DataGrid1.Columns(1).Text = txtFields(34).Text
    ADOItemRS.Update
    grdOnAddNew = False
    ProductionItem True
Else
    ADOItemRS.Update
    grdOnAddNew = False
    ProductionItem True
End If
DataGrid1_AfterColEdit 0
Exit Sub
Error_ButtClick:
    MsgBox "Please click the Table box before clicking the button"
End Sub

Private Sub INV_ITEM()
   AllLookup.GetWhichTable 1750, "SELECT [INV ITEM Id], [INV ITEM Description]," & _
   "[INV ITEM Unit],[INV ITEM Price], [INV ITEM Sales Account], [INV ITEM Qty On Hand], " & _
   "[INV ITEM Qty On Order], [INV ITEM Taxable YN],[INV ITEM Last Cost] FROM [INV Items] " & _
   "WHERE [INV ITEM Type] <> 'Assembly'", "Product", _
   "Item ID//Item Description//Unit//Price//Sales Account//Qty On Hand//Qty On Order//Taxable//Cost", db
End Sub

Private Sub DataGrid1_Error(ByVal DataError As Integer, Response As Integer)
    If DataGridKnownError(DataError) Then
        Response = 0
    End If
End Sub

Private Sub DataGrid1_GotFocus()
Dim CreateOrder As Integer
    If mbAddNewFlag = True Then
        'cmdAdd.SetFocus
        CreateOrder = MsgBox("This Request will save the data to the database? Are sure to continue", vbYesNo, "Save")
        If CreateOrder = vbNo Then Exit Sub
        cmdAdd_Click
    End If
End Sub

Private Sub DataGrid1_OnAddNew()
    grdOnAddNew = True
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If ADOItemRS.BOF Or ADOItemRS.EOF Then Exit Sub
        If DataGrid1.col > 2 And DataGrid1.Row > -1 Then
            If DataGrid1.Columns(2) = "" Then
                MsgBox "You must select Item ID first before continue", vbInformation, "Error Selection"
                GoTo Damn_Attempt
            End If
        End If
        ItemsCost
Select Case DataGrid1.col
  Case 3
     DataGrid1.AllowUpdate = True
  Case Else
     DataGrid1.AllowUpdate = False
  End Select
Exit Sub
Damn_Attempt:
     DataGrid1.AllowUpdate = False
     DataGrid1.col = 0
End Sub

Private Sub ItemsCost()
Dim ADOItemsCostRS As ADODB.Recordset
Dim CostValue As Double

CostValue = 0

Set ADOItemsCostRS = New ADODB.Recordset
ADOItemsCostRS.Open "SELECT [INV KIT ID],[INV KIT Cost],[INV KIT Qty] FROM [INV Kit Items 2] " & _
"WHERE [INV KIT Item ID]='" & txtFields(34) & "'", db, adOpenKeyset, adLockOptimistic, adCmdText

With ADOItemsCostRS
If .RecordCount > 0 Then
    .MoveFirst
    Do While Not .EOF
        CostValue = CostValue + (![INV KIT Cost] * ![INV KIT Qty])
        .MoveNext
    Loop
End If
ADOprimaryrs![INV ITEM Last Cost] = CostValue
ADOprimaryrs.Update
End With

ADOItemsCostRS.Close
Set ADOItemsCostRS = Nothing
End Sub

Private Sub Form_Load()
ShowStatus True
'On Error GoTo FormErr
    
    Me.Height = 7635
    Me.Width = 10200
    
    Set db = New ADODB.Connection
    db.CursorLocation = adUseClient
    db.Open gblADOProvider
    
  'use to identify whether the ADOsecondaryRS is loaded
  SecondaryLoad = False
  ThirdLoad = False
        '[INV ITEM Unit]
        '[INV ITEM Price]
  Dim sql As String
  sql = "select [INV ITEM Id],[INV ITEM Description],[INV ITEM Type],[INV ITEM Unit]," & _
    "[INV ITEM Category],[INV ITEM Commissionable YN],[INV ITEM Taxable YN],[INV ITEM Price]," & _
    "[INV ITEM Vendor ID],[INV ITEM Vendors Number],[INV ITEM Sales Account]," & _
    "[INV ITEM Inventory Account],[INV ITEM Cost of Sales Account]," & _
    "[INV ITEM Inactive YN],[INV ITEM Qty On Hand],[INV ITEM Qty On Order]," & _
    "[INV ITEM Reorder Qty],[INV ITEM Costing Method],[INV ITEM Standard Cost]," & _
    "[INV ITEM Last Cost],[INV ITEM Average Cost] from [INV Items] " & _
    "Order by [INV ITEM Id]"
  Set ADOprimaryrs = New ADODB.Recordset
  ADOprimaryrs.Open sql, db, adOpenStatic, adLockOptimistic
  
  Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtFields
    If Trim(oText.DataField) <> "" Then
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
  'Bind the Datacombos to the data provider
  For Each oCombo In Me.cbfields
    Set oCombo.DataSource = ADOprimaryrs
  Next
  
  loadCombo
  
  'Lock these fields to avoid invalid entries
  'txtFields(6).Locked = True
  'txtFields(8).Locked = True
  'txtFields(9).Locked = True
  'txtFields(10).Locked = True
  
  If CheckNewDB(ADOprimaryrs, "Items") = True Then
    cmdAdd_Click
  Else
    ItemPrice "' AND [INV BREAK Default]=TRUE"
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
  
  Me.Width = 10470
  Me.Height = 7905
  
SkipResize:
  frPrimary.Left = (Me.ScaleWidth - frPrimary.Width) / 2
  Picture1.Left = frPrimary.Left
  'lblTop.Width = frPrimary.Width
  frPrimary.Top = (Me.ScaleHeight - frPrimary.Height - picButtons.Height - picStatBox.Height) / 2 + 230
  
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
  If txtPrice(0) = "" And Not mbAddNewFlag Then
    MsgBox "You have to set the item price, skipping this step might generate an error in the future.", vbCritical, "Error"
    Exit Sub
  End If
      
  EndLoad db, ADOsecondaryRS, "Items"
  Set grdDataGrid1.DataSource = Nothing
 If frPricing.Visible = True Then
    cmdBackPricing_Click
  End If
  ShowStatus True
      If ADOprimaryrs.RecordCount > 0 Then
        If ADOprimaryrs.EditMode <> 0 Then
          ADOprimaryrs.CancelUpdate
        End If
      End If
      If txtPrice(0) <> "" Then
        ADOsecondaryRS.CancelUpdate
        ADOsecondaryRS.Close
        Set ADOsecondaryRS = Nothing
      End If
        If ADOItemRS Is Nothing Then
        Else
            ADOItemRS.Close
            Set ADOItemRS = Nothing
        End If
      ADOprimaryrs.Close
      Set ADOprimaryrs = Nothing
      db.Close
      Set db = Nothing
  Set frm_SYS_Setup_Items = Nothing
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
  If SecondaryLoad = True Then ItemPrice "' AND [INV BREAK Default]=TRUE"
  CurrItem = ADOprimaryrs![INV ITEM Id] & ""
  txtFields(34).Text = CurrItem
  If IsNull(ADOprimaryrs![INV ITEM Type]) = False Then
    If ADOprimaryrs![INV ITEM Type] = "Assembly" Then
        Picture3.ZOrder 0
        ProductionItem True
    Else
        Picture2.ZOrder 0
        ProductionItem False
    End If
  Else
    Picture2.ZOrder 0
  End If
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
    If cmdAdd.Caption = "&Add" Then
        If Not (.BOF And .EOF) And Not mbAddNewFlag Then
            mvBookMark = .Bookmark
        End If
        mbAddNewFlag = True
        .AddNew
        lblStatus.Caption = "Add record"
        txtFields(0).Enabled = True
        'txtfields(0).SetFocus
        cmdAdd.Caption = "&Cancel"
        cmdUpdate.Enabled = True
    Else
        mbAddNewFlag = False
        .CancelUpdate
        txtFields(0).Enabled = False
        cmdAdd.Caption = "&Add"
        If .RecordCount > 0 Then
            If mvBookMark > 0 Then
                .Bookmark = mvBookMark
            Else
                .MoveFirst
            End If
        End If
        'ItemPrice "' AND [INV BREAK Default]=TRUE"
    End If
    
    'set to controls appropriately
    cmdDelete.Enabled = Not mbAddNewFlag
    cmdRefresh.Enabled = Not mbAddNewFlag
  End With
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

Private Sub cmdUpdate_Click()
    On Error GoTo UpdateErr

    With ADOprimaryrs
        If .RecordCount = 0 Then Exit Sub 'no records to update
        If Trim(txtFields(0).Text) <> "" Then
        Dim oTxt As TextBox
          For Each oTxt In Me.txtFields
            If oTxt.Text = "" Then
              If ADOprimaryrs("" & oTxt.DataField & "").Type = 203 Or ADOprimaryrs("" & oTxt.DataField & "").Type = 202 Then oTxt.Text = " "
            End If
          Next
        Else
            MsgBox lblLabels(0) & " must be filled. Please try again before Update.", vbInformation, "Information"
            Exit Sub
        End If
        ![INV ITEM Inactive YN] = True
        .Update
        
        Dim TempItemID As String 'INV ITEM Id
        TempItemID = txtFields(34).Text
        If mbAddNewFlag Then 'requery to get default value assigned by database
            .Requery
            .Find "[INV ITEM Id]='" & TempItemID & "'"
            mbAddNewFlag = False
        End If
        
        'reenable the necessary buttons
        cmdAdd.Caption = "&Add"
        txtFields(0).Enabled = False
        cmdDelete.Enabled = True
        cmdRefresh.Enabled = True
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
  cmdAdd.Enabled = bVal
  cmdUpdate.Enabled = bVal
  'cmdCancel.Enabled = Not bVal
  cmdDelete.Enabled = bVal
  'cmdClose.Enabled = bVal
  cmdRefresh.Enabled = bVal
  cmdNext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdLast.Enabled = bVal
  cmdPrevious.Enabled = bVal
End Sub

Private Sub loadCombo(Optional Index As Integer)
    Index = IIf(Index > 0, Index, 0)
    Select Case Index
    Case 0
        ComboInit cbfields(1), lblfields(1), "select [Type] from [LIST Item Types]"
        ComboInit cbfields(2), lblfields(2), "select [Category] from [LIST Item Categories]"
    Case 1
        ComboInit cbfields(1), lblfields(1), "select [Type] from [LIST Item Types]"
    Case 2
        ComboInit cbfields(2), lblfields(2), "select [Category] from [LIST Item Categories]"
    End Select
        
End Sub

Private Sub grdDatagrid1_BeforeDelete(Cancel As Integer)
    Dim DeleteCration As Integer
    
    DeleteCration = MsgBox("Attempting to delete the data. " & vbCr & "Are you sure?", vbYesNo, "Deleting Confirmation")
    If DeleteCration = vbNo Then Cancel = 1
End Sub

Private Sub grdDatagrid3_HeadClick(ByVal ColIndex As Integer)
    ADOfourthRS.Close
    Set ADOfourthRS = New ADODB.Recordset
    'MsgBox grdDatagrid3.Columns(ColIndex).DataField
    ADOfourthRS.Open "select [INV ITEM Id],[INV ITEM Description],[INV ITEM Vendor ID],[INV ITEM Last Cost],[INV ITEM Price],[INV ITEM Qty On Hand],[INV ITEM Qty On Order],[INV ITEM Notes] from [INV Items] Order by [" & grdDataGrid3.Columns(ColIndex).DataField & "]", db, adOpenKeyset, adLockOptimistic, adCmdText
    Set grdDataGrid3.DataSource = ADOfourthRS
    ADOfourthRS.Requery
End Sub

Private Sub ProductionItem(CallDB As Boolean)

    Set DataGrid1.DataSource = Nothing
    If ADOItemRS Is Nothing Then
    Else
        ADOItemRS.Close
        Set ADOItemRS = Nothing
    End If
    
If CallDB = True Then
    Set ADOItemRS = New ADODB.Recordset
    
    ADOItemRS.Open "Select * from [INV Kit Items 2] WHERE [INV KIT Item ID]='" & _
    txtFields(34) & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
    'MsgBox ADOItemRS.RecordCount & "   " & txtFields(34)
    Set DataGrid1.DataSource = ADOItemRS
End If
End Sub


Private Sub optPriceList_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    grdDataGrid3.Visible = False
If optPriceList(Index).Value = 1 Then
    grdDataGrid3.Columns(Index).Visible = True
Else
    grdDataGrid3.Columns(Index).Visible = False
End If
grdDataGrid3.Visible = True
End Sub

Private Sub txtfields_GotFocus(Index As Integer)
    TxtGotFocus txtFields(Index)
End Sub

Private Sub txtfields_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
Case 15, 16
    keyResponse = CtrlValidate(KeyAscii, "0123456789.")
    If keyResponse = True Then
    Else
       KeyAscii = 0
    End If
Case 8, 9, 10
    keyResponse = CtrlValidate(KeyAscii, "0123456789")
    If keyResponse = True Then
    Else
       KeyAscii = 0
    End If
End Select
End Sub

Private Sub txtFields_LostFocus(Index As Integer)
Select Case Index
Case 15
    txtFields(Index) = Val(txtFields(Index))
Case 16
    If txtFields(Index) <> "" Then txtFields(Index) = FormatCurr(txtFields(Index))
Case 8, 9, 10
    'MsgBox BankAcct35 & "   " & txtFields(Index)
    If txtFields(Index) = "" Then Exit Sub
    If IsNumeric(txtFields(Index).Text) Then
        CheckDocument "SELECT [GL COA Account No] FROM [GL Chart Of Accounts] WHERE [GL COA Account No]='" & txtFields(Index).Text & "'", db, False, txtFields(Index)
    Else
        MsgBox "Only numeric character is accepted", vbInformation, "Information"
        txtFields(Index) = ""
    End If
    'BankAcct35 = txtFields(Index)
Case 34
    Dim txtItemID As String
    txtItemID = txtFields(34)
    If txtFields(34) = "" And mbAddNewFlag = False Then
        txtFields(34) = CurrItem
    ElseIf txtFields(34) <> "" And mbAddNewFlag = True Then
        If CheckDocument("SELECT [INV ITEM Id] FROM [INV Items] WHERE [INV ITEM Id]='" & txtItemID & "'", db, False) = False Then
            MsgBox txtItemID & " is already exist", vbInformation, "Information"
            txtFields(34) = ""
            txtFields(0) = txtFields(34)
            Exit Sub
        Else
            txtFields(0) = txtItemID
        End If
    End If
    
    If txtFields(34) = CurrItem Then Exit Sub
    
    With ADOprimaryrs
      If .RecordCount > 0 And mbAddNewFlag = False Then
        mvBookMark = .Bookmark
        .MoveFirst
        .Find "[INV ITEM Id]='" & txtItemID & "'"
        If .EOF Then
            Dim Response As Integer
            Response = MsgBox(txtItemID & " is a new input. Would you like to add it into the database", vbYesNo, "Information")
            If Response = vbYes Then
                mbAddNewFlag = True
                cmdAdd_Click
                txtFields(34) = txtItemID
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


Private Sub txtPrice_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
Case 4
    grdDataGrid1.AllowUpdate = True
    keyResponse = CtrlValidate(KeyAscii, "0123456789.")
    If keyResponse = True Then
    Else
       KeyAscii = 0
    End If
Case 6
    grdDataGrid1.AllowUpdate = True
    keyResponse = CtrlValidate(KeyAscii, "0123456789")
    If keyResponse = True Then
    Else
       KeyAscii = 0
    End If
Case 7
    grdDataGrid1.AllowUpdate = True
Case Else
    grdDataGrid1.AllowUpdate = False
End Select
End Sub

Private Sub txtPrice_LostFocus(Index As Integer)

Select Case Index
Case 6
    txtPrice(Index) = Val(txtPrice(Index))
    txtPrice(5) = FormatCurr(txtPrice(6) * txtPrice(4))
    'ADOsecondaryRS![INV BREAK Price] = FormatCurr(txtPrice(6) * txtPrice(4))
Case 4
    txtPrice(Index) = FormatCurr(txtPrice(Index))
    txtPrice(5) = txtPrice(6) * txtPrice(4)
    'ADOsecondaryRS![INV BREAK Price] = txtPrice(6) * txtPrice(4)
End Select
    If ADOsecondaryRS.EditMode <> adEditNone Then
        'Set grdDatagrid1.DataSource = Nothing
        ADOsecondaryRS![INV BREAK Price] = txtPrice(5) & ""
        grdDataGrid1.SetFocus
        SendKeys ("{Left}")
        ADOsecondaryRS.Update
        'ADOsecondaryRS.Update
    End If
    
    If txtPrice(5) = 0 Then
        MsgBox "Item Pricing cannot be zero", vbInformation, "Information"
    End If
    grdDataGrid1.AllowUpdate = False
End Sub

