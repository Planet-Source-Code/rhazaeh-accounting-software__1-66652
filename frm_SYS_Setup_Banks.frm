VERSION 5.00
Begin VB.Form frm_SYS_Setup_Banks 
   Caption         =   "Bank Setup"
   ClientHeight    =   7080
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   12315
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7080
   ScaleWidth      =   12315
   Begin VB.Frame frPrimary 
      Height          =   6255
      Left            =   0
      TabIndex        =   35
      Top             =   480
      Width           =   12255
      Begin VB.Frame Frame1 
         Height          =   2775
         Left            =   120
         TabIndex        =   63
         Top             =   720
         Width           =   6135
         Begin VB.CommandButton btGLFinanceCharge 
            Height          =   285
            Left            =   5160
            Picture         =   "frm_SYS_Setup_Banks.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   2280
            Width           =   375
         End
         Begin VB.CommandButton btGLServiceCharges 
            Height          =   285
            Left            =   5160
            Picture         =   "frm_SYS_Setup_Banks.frx":014A
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   1920
            Width           =   375
         End
         Begin VB.CommandButton btGLOtherChanges 
            Height          =   285
            Left            =   5160
            Picture         =   "frm_SYS_Setup_Banks.frx":0294
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   1560
            Width           =   375
         End
         Begin VB.CommandButton btGLMiscWithdrawal 
            Height          =   285
            Left            =   5160
            Picture         =   "frm_SYS_Setup_Banks.frx":03DE
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   1200
            Width           =   375
         End
         Begin VB.CommandButton btGLMiscDeposit 
            Height          =   285
            Left            =   5160
            Picture         =   "frm_SYS_Setup_Banks.frx":0528
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   840
            Width           =   375
         End
         Begin VB.CommandButton btGLInterestEarned 
            Height          =   285
            Left            =   5160
            Picture         =   "frm_SYS_Setup_Banks.frx":0672
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox txtFields 
            DataField       =   "SYS COM Bank Interest Earned Acct"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   0
            Left            =   3240
            Locked          =   -1  'True
            TabIndex        =   69
            Top             =   480
            Width           =   1935
         End
         Begin VB.TextBox txtFields 
            DataField       =   "SYS COM Bank Misc Deposit Acct"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   1
            Left            =   3240
            Locked          =   -1  'True
            TabIndex        =   68
            Top             =   840
            Width           =   1935
         End
         Begin VB.TextBox txtFields 
            DataField       =   "SYS COM Bank Misc Withdrawl Acct"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   2
            Left            =   3240
            Locked          =   -1  'True
            TabIndex        =   67
            Top             =   1200
            Width           =   1935
         End
         Begin VB.TextBox txtFields 
            DataField       =   "SYS COM Bank Other Charges Acct"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   3
            Left            =   3240
            Locked          =   -1  'True
            TabIndex        =   66
            Top             =   1560
            Width           =   1935
         End
         Begin VB.TextBox txtFields 
            DataField       =   "SYS COM Bank Service Charges Acct"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   4
            Left            =   3240
            Locked          =   -1  'True
            TabIndex        =   65
            Top             =   1920
            Width           =   1935
         End
         Begin VB.TextBox txtFields 
            DataField       =   "SYS COM Finance Charge Acct"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   5
            Left            =   3240
            Locked          =   -1  'True
            TabIndex        =   64
            Top             =   2280
            Width           =   1935
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "GL Bank Interest Earned Account:  "
            Height          =   255
            Index           =   1
            Left            =   480
            TabIndex        =   75
            Top             =   480
            Width           =   2775
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "GL Bank Misc Deposit Account:  "
            Height          =   255
            Index           =   2
            Left            =   480
            TabIndex        =   74
            Top             =   840
            Width           =   2775
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "GL Bank Misc Withdrawal Account:  "
            Height          =   255
            Index           =   3
            Left            =   480
            TabIndex        =   73
            Top             =   1200
            Width           =   2775
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "GL Bank Other Charges Acct:  "
            Height          =   255
            Index           =   4
            Left            =   480
            TabIndex        =   72
            Top             =   1560
            Width           =   2775
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "GL Bank Service Charges Acct:  "
            Height          =   255
            Index           =   5
            Left            =   480
            TabIndex        =   71
            Top             =   1920
            Width           =   2775
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "GL Bank Finance Charge Acct:  "
            Height          =   255
            Index           =   6
            Left            =   480
            TabIndex        =   70
            Top             =   2280
            Width           =   2775
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2175
         Left            =   120
         TabIndex        =   55
         Top             =   3960
         Width           =   6135
         Begin VB.TextBox txtFields 
            DataField       =   "SYS COM Annual Charge"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   6
            Left            =   3120
            TabIndex        =   7
            Top             =   240
            Width           =   1815
         End
         Begin VB.TextBox txtFields 
            DataField       =   "SYS COM Minimum Balance"
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
            Index           =   7
            Left            =   3120
            TabIndex        =   8
            Top             =   600
            Width           =   2175
         End
         Begin VB.TextBox txtFields 
            DataField       =   "SYS COM Minimum Finance Charge"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   8
            Left            =   3120
            TabIndex        =   9
            Top             =   960
            Width           =   1815
         End
         Begin VB.TextBox txtFields 
            DataField       =   "SYS COM Monthly Charge"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   9
            Left            =   3120
            TabIndex        =   10
            Top             =   1320
            Width           =   1815
         End
         Begin VB.CheckBox chkFields 
            Alignment       =   1  'Right Justify
            Caption         =   "Interest (Y/N)"
            DataField       =   "SYS COM Interest YN"
            DataSource      =   "adoPrimaryRS"
            Height          =   255
            Index           =   11
            Left            =   3600
            TabIndex        =   12
            Top             =   1680
            Width           =   1695
         End
         Begin VB.CheckBox chkFields 
            Alignment       =   1  'Right Justify
            Caption         =   "Finance Charge (Y/N)"
            DataField       =   "SYS COM Finance Charges YN"
            DataSource      =   "adoPrimaryRS"
            Height          =   255
            Index           =   10
            Left            =   1240
            TabIndex        =   11
            Top             =   1680
            Width           =   2055
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Bank Annual Charge:  "
            Height          =   255
            Index           =   0
            Left            =   720
            TabIndex        =   62
            Top             =   240
            Width           =   2415
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Bank Minimum Balance:  "
            Height          =   255
            Index           =   9
            Left            =   720
            TabIndex        =   61
            Top             =   600
            Width           =   2415
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Bank Minimum Finance Charge:  "
            Height          =   255
            Index           =   10
            Left            =   720
            TabIndex        =   60
            Top             =   960
            Width           =   2415
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Bank Monthly Charge:  "
            Height          =   255
            Index           =   11
            Left            =   720
            TabIndex        =   59
            Top             =   1320
            Width           =   2415
         End
         Begin VB.Label lblPct 
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
            Left            =   5040
            TabIndex        =   58
            Top             =   240
            Width           =   210
         End
         Begin VB.Label lblPct 
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
            Index           =   1
            Left            =   5040
            TabIndex        =   57
            Top             =   960
            Width           =   210
         End
         Begin VB.Label lblPct 
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
            Index           =   2
            Left            =   5040
            TabIndex        =   56
            Top             =   1320
            Width           =   210
         End
      End
      Begin VB.ComboBox cbBank 
         Height          =   315
         Left            =   6480
         TabIndex        =   13
         Top             =   720
         Width           =   1815
      End
      Begin VB.CommandButton cmdBank 
         Height          =   280
         Left            =   8280
         Picture         =   "frm_SYS_Setup_Banks.frx":07BC
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Update the Ship Via"
         Top             =   720
         Width           =   375
      End
      Begin VB.Frame Frame3 
         Height          =   5415
         Left            =   6360
         TabIndex        =   37
         Top             =   720
         Width           =   5775
         Begin VB.TextBox txtBank 
            DataField       =   "BANK ACCT Number"
            Height          =   285
            Index           =   1
            Left            =   1680
            TabIndex        =   15
            Top             =   1080
            Width           =   1575
         End
         Begin VB.TextBox txtBank 
            DataField       =   "BANK ACCT Name"
            Height          =   285
            Index           =   2
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   1440
            Width           =   3975
         End
         Begin VB.TextBox txtBank 
            DataField       =   "BANK ACCT Address 1"
            Height          =   285
            Index           =   3
            Left            =   1680
            TabIndex        =   17
            Top             =   1800
            Width           =   3975
         End
         Begin VB.TextBox txtBank 
            DataField       =   "BANK ACCT Address 2"
            Height          =   285
            Index           =   4
            Left            =   1680
            TabIndex        =   18
            Top             =   2160
            Width           =   3975
         End
         Begin VB.TextBox txtBank 
            DataField       =   "BANK ACCT City"
            Height          =   285
            Index           =   5
            Left            =   1680
            TabIndex        =   19
            Top             =   2520
            Width           =   975
         End
         Begin VB.TextBox txtBank 
            DataField       =   "BANK ACCT State"
            Height          =   285
            Index           =   6
            Left            =   3360
            TabIndex        =   20
            Top             =   2520
            Width           =   735
         End
         Begin VB.TextBox txtBank 
            DataField       =   "BANK ACCT Postal"
            Height          =   285
            Index           =   7
            Left            =   4800
            TabIndex        =   21
            Top             =   2520
            Width           =   855
         End
         Begin VB.TextBox txtBank 
            DataField       =   "BANK ACCT Country"
            Height          =   285
            Index           =   8
            Left            =   1680
            TabIndex        =   22
            Top             =   2880
            Width           =   1335
         End
         Begin VB.TextBox txtBank 
            DataField       =   "BANK ACCT Phone"
            Height          =   285
            Index           =   9
            Left            =   1680
            TabIndex        =   23
            Top             =   3240
            Width           =   1335
         End
         Begin VB.TextBox txtBank 
            DataField       =   "BANK ACCT Fax"
            Height          =   285
            Index           =   10
            Left            =   4320
            TabIndex        =   24
            Top             =   3240
            Width           =   1335
         End
         Begin VB.TextBox txtBank 
            DataField       =   "BANK ACCT Officer"
            Height          =   285
            Index           =   11
            Left            =   1680
            TabIndex        =   25
            Top             =   3600
            Width           =   3975
         End
         Begin VB.TextBox txtBank 
            DataField       =   "BANK ACCT Next Check No"
            Height          =   285
            Index           =   12
            Left            =   1680
            TabIndex        =   26
            Top             =   3960
            Width           =   1335
         End
         Begin VB.TextBox txtBank 
            DataField       =   "BANK ACCT Next Deposit No"
            Height          =   285
            Index           =   13
            Left            =   1680
            TabIndex        =   27
            Top             =   4320
            Width           =   1335
         End
         Begin VB.TextBox txtBank 
            DataField       =   "BANK ACCT Notes"
            Height          =   1125
            Index           =   14
            Left            =   3240
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   30
            Top             =   4200
            Width           =   2415
         End
         Begin VB.TextBox txtBank 
            DataField       =   "BANK ACCT Balance"
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
            Index           =   15
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   28
            Top             =   4680
            Width           =   1335
         End
         Begin VB.TextBox txtBank 
            DataField       =   "BANK ACCT Unposted Deposits"
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
            Index           =   16
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   29
            Top             =   5040
            Width           =   1335
         End
         Begin VB.CommandButton cmdBankUpdate 
            Caption         =   "Update"
            Enabled         =   0   'False
            Height          =   855
            Left            =   4680
            Picture         =   "frm_SYS_Setup_Banks.frx":0AC6
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox txtBank 
            DataField       =   "BANK ACCT ID"
            Height          =   285
            Index           =   0
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label lblBank 
            Alignment       =   1  'Right Justify
            Caption         =   "Account Number:"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   54
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label lblBank 
            Alignment       =   1  'Right Justify
            Caption         =   "Bank Name:"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   53
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label lblBank 
            Alignment       =   1  'Right Justify
            Caption         =   "Bank Address:"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   52
            Top             =   1800
            Width           =   1335
         End
         Begin VB.Label lblBank 
            Alignment       =   1  'Right Justify
            Caption         =   "City:"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   51
            Top             =   2520
            Width           =   1455
         End
         Begin VB.Label lblBank 
            Alignment       =   1  'Right Justify
            Caption         =   "State:"
            Height          =   255
            Index           =   6
            Left            =   2760
            TabIndex        =   50
            Top             =   2520
            Width           =   495
         End
         Begin VB.Label lblBank 
            Alignment       =   1  'Right Justify
            Caption         =   "Postal:"
            Height          =   255
            Index           =   7
            Left            =   4200
            TabIndex        =   49
            Top             =   2520
            Width           =   495
         End
         Begin VB.Label lblBank 
            Alignment       =   1  'Right Justify
            Caption         =   "Country:"
            Height          =   255
            Index           =   8
            Left            =   240
            TabIndex        =   48
            Top             =   2880
            Width           =   1335
         End
         Begin VB.Label lblBank 
            Alignment       =   1  'Right Justify
            Caption         =   "Phone Number:"
            Height          =   255
            Index           =   9
            Left            =   240
            TabIndex        =   47
            Top             =   3240
            Width           =   1335
         End
         Begin VB.Label lblBank 
            Alignment       =   1  'Right Justify
            Caption         =   "Fax Number:"
            Height          =   255
            Index           =   10
            Left            =   3240
            TabIndex        =   46
            Top             =   3240
            Width           =   975
         End
         Begin VB.Label lblBank 
            Alignment       =   1  'Right Justify
            Caption         =   "Officer Name:"
            Height          =   255
            Index           =   11
            Left            =   240
            TabIndex        =   45
            Top             =   3600
            Width           =   1335
         End
         Begin VB.Label lblBank 
            Alignment       =   1  'Right Justify
            Caption         =   "Next Check No:"
            Height          =   255
            Index           =   12
            Left            =   240
            TabIndex        =   44
            Top             =   3960
            Width           =   1335
         End
         Begin VB.Label lblBank 
            Alignment       =   1  'Right Justify
            Caption         =   "Next Deposit No:"
            Height          =   255
            Index           =   13
            Left            =   240
            TabIndex        =   43
            Top             =   4320
            Width           =   1335
         End
         Begin VB.Label lblBank 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Notes"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   14
            Left            =   3240
            TabIndex        =   42
            Top             =   3960
            Width           =   2415
         End
         Begin VB.Label lblBank 
            Alignment       =   1  'Right Justify
            Caption         =   "Bank Balance:"
            Height          =   255
            Index           =   15
            Left            =   240
            TabIndex        =   41
            Top             =   4680
            Width           =   1335
         End
         Begin VB.Label lblBank 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Unposted Deposits:"
            Height          =   255
            Index           =   16
            Left            =   120
            TabIndex        =   40
            Top             =   5040
            Width           =   1455
         End
         Begin VB.Label lblBankTemp 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Bank"
            Height          =   195
            Left            =   1080
            TabIndex        =   39
            Top             =   0
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label lblBank 
            Alignment       =   1  'Right Justify
            Caption         =   "COA Number:"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   38
            Top             =   720
            Width           =   1335
         End
      End
      Begin VB.Label lblAcct 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "General Ledger Bank Setup"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   78
         Top             =   360
         Width           =   6135
      End
      Begin VB.Label lblAcct 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "General Ledger Bank Setup"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   77
         Top             =   3600
         Width           =   6135
      End
      Begin VB.Label lblAcct 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bank Particulars"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   1
         Left            =   6360
         TabIndex        =   76
         Top             =   360
         Width           =   5775
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
      ScaleWidth      =   12315
      TabIndex        =   0
      Top             =   6780
      Width           =   12315
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   300
         Left            =   2160
         TabIndex        =   34
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   300
         Left            =   1080
         TabIndex        =   33
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   300
         Left            =   0
         TabIndex        =   32
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.Label lblTop 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bank Setup"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   79
      Top             =   120
      Width           =   5775
   End
End
Attribute VB_Name = "frm_SYS_Setup_Banks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ADOprimaryrs As ADODB.Recordset
Attribute ADOprimaryrs.VB_VarHelpID = -1
Dim ADOsecondaryRS As ADODB.Recordset
Dim db As ADODB.Connection

Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean

Private Sub btGLFinanceCharge_Click()
    Dim SQLstatement As String
    Dim No As Integer
    Dim ghead As String
    Dim fhead As String

    No = 15
    SQLstatement = "select [GL COA Account No], [GL COA Account Name]" & _
                    "from [GL Chart Of Accounts]"
    ghead = "Account Description"
    fhead = "Account No//Description"
    
    AllLookup.GetWhichTable No, SQLstatement, ghead, fhead, db
    'AllLookup.Show vbModal
    txtFields(5).SetFocus
End Sub

Private Sub btGLInterestEarned_Click()
    Dim SQLstatement As String
    Dim No As Integer
    Dim ghead As String
    Dim fhead As String

    No = 10
    SQLstatement = "select [GL COA Account No], [GL COA Account Name]" & _
                    "from [GL Chart Of Accounts]"
    ghead = "Account Description"
    fhead = "Account No//Description"
    
    AllLookup.GetWhichTable No, SQLstatement, ghead, fhead, db
    'AllLookup.Show vbModal
    txtFields(0).SetFocus
End Sub

Private Sub btGLMiscDeposit_Click()
    Dim SQLstatement As String
    Dim No As Integer
    Dim ghead As String
    Dim fhead As String

    No = 11
    SQLstatement = "select [GL COA Account No], [GL COA Account Name]" & _
                    "from [GL Chart Of Accounts]"
    ghead = "Account Description"
    fhead = "Account No//Description"
    
    AllLookup.GetWhichTable No, SQLstatement, ghead, fhead, db
    'AllLookup.Show vbModal
    txtFields(1).SetFocus
End Sub

Private Sub btGLMiscWithdrawal_Click()
    Dim SQLstatement As String
    Dim No As Integer
    Dim ghead As String
    Dim fhead As String

    No = 12
    SQLstatement = "select [GL COA Account No], [GL COA Account Name]" & _
                    "from [GL Chart Of Accounts]"
    ghead = "Account Description"
    fhead = "Account No//Description"
    
    AllLookup.GetWhichTable No, SQLstatement, ghead, fhead, db
    'AllLookup.Show vbModal
    txtFields(2).SetFocus
End Sub

Private Sub btGLOtherChanges_Click()
    Dim SQLstatement As String
    Dim No As Integer
    Dim ghead As String
    Dim fhead As String

    No = 13
    SQLstatement = "select [GL COA Account No], [GL COA Account Name]" & _
                    "from [GL Chart Of Accounts]"
    ghead = "Account Description"
    fhead = "Account No//Description"
    
    AllLookup.GetWhichTable No, SQLstatement, ghead, fhead, db
    'AllLookup.Show vbModal
    txtFields(3).SetFocus
End Sub

Private Sub btGLServiceCharges_Click()
    Dim SQLstatement As String
    Dim No As Integer
    Dim ghead As String
    Dim fhead As String

    No = 14
    SQLstatement = "select [GL COA Account No], [GL COA Account Name]" & _
                    "from [GL Chart Of Accounts]"
    ghead = "Account Description"
    fhead = "Account No//Description"
    
    AllLookup.GetWhichTable No, SQLstatement, ghead, fhead, db
    'AllLookup.Show vbModal
    txtFields(4).SetFocus
End Sub

Private Sub cbBank_Click()
  If cbBank.List(0) <> "" Then
    Dim oTxtBank As TextBox
    For Each oTxtBank In Me.txtBank
        Set oTxtBank.DataSource = Nothing
    Next
    ADOsecondaryRS.Close
    Set ADOsecondaryRS = Nothing
  End If
  
  Secondary cbBank.Text
    If txtBank(0) = "" Then
        ADOsecondaryRS.AddNew
            txtBank(0) = LookRecord("[GL COA Account No]", "[GL Chart Of Accounts]", db, "[GL COA Account Name] = '" & cbBank & "'")
            txtBank(1) = 0
            txtBank(2) = cbBank.Text
        ADOsecondaryRS.Update
    End If
    'txtBank(2)
    'If oTxtBank.Text <> "" Then lblAcct(oTxtBank.Index) = LookRecord("[GL COA Account Name]", "[GL Chart Of Accounts]", "[GL COA Account No] = '" & txtFields(oTxtBank.Index) & "'")
End Sub

Private Sub cbBank_KeyPress(KeyAscii As Integer)
    keyResponse = CtrlValidate(KeyAscii, "")
    If keyResponse = True Then
    Else
       KeyAscii = 0
    End If
End Sub

Private Sub cbBank_LostFocus()
   If CheckCombo(cbBank) Then
        MsgBox "There is no such selection", vbInformation, "Information"
   End If
End Sub

Private Sub cmdBank_Click()
    loadCombo
  If cbBank.List(0) <> "" Then
    Dim oTxtBank As TextBox
    For Each oTxtBank In Me.txtBank
        Set oTxtBank.DataSource = Nothing
    Next
    
    ADOsecondaryRS.CancelUpdate
    ADOsecondaryRS.Close
    Set ADOsecondaryRS = Nothing
  End If
    Secondary
End Sub

Private Sub loadCombo()
    ComboInit cbBank, lblBank(0), "SELECT [GL COA Account Name] FROM [GL Chart Of Accounts] WHERE [GL COA Asset Type]='Cash'"
End Sub

Private Sub cmdBankUpdate_Click()
  On Error GoTo UpdateErr
  
  Dim oTxt As TextBox
   For Each oTxt In Me.txtBank
      If oTxt.Text = "" Then
         If ADOsecondaryRS("" & oTxt.DataField & "").Type = 203 Or ADOsecondaryRS("" & oTxt.DataField & "").Type = 202 Then oTxt.Text = " "
      End If
   Next
  With ADOsecondaryRS
    .Update
  End With
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub Form_Load()
ShowStatus True
On Error GoTo FormErr
  Set db = New ADODB.Connection
  db.CursorLocation = adUseClient
  db.Open gblADOProvider

  Set ADOprimaryrs = New ADODB.Recordset
  ADOprimaryrs.Open "select [SYS COM Annual Charge],[SYS COM Bank Interest Earned Acct]," & _
    "[SYS COM Bank Misc Deposit Acct],[SYS COM Bank Misc Withdrawl Acct]," & _
    "[SYS COM Bank Other Charges Acct],[SYS COM Bank Service Charges Acct]," & _
    "[SYS COM Finance Charge Acct],[SYS COM Finance Charges YN],[SYS COM Interest YN]," & _
    "[SYS COM Minimum Balance],[SYS COM Minimum Finance Charge],[SYS COM Monthly Charge] " & _
    "from [SYS Company]", db, adOpenStatic, adLockOptimistic

  Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtFields
    Set oText.DataSource = ADOprimaryrs
    If ADOprimaryrs("" & oText.DataField & "").Type = 202 Then oText.MaxLength = ADOprimaryrs("" & oText.DataField & "").DefinedSize
  Next

  Dim oCheck As CheckBox
    'Bind the Check boxes to the data provider
  For Each oCheck In Me.chkFields
    Set oCheck.DataSource = ADOprimaryrs
  Next
  
  If CheckNewDB(ADOprimaryrs, "Banks") = True Then
    ADOprimaryrs.AddNew
  End If
  
  loadCombo
  Secondary cbBank.Text
  
  GetTextColor Me
  mbDataChanged = False
ShowStatus False
Exit Sub
FormErr:
  MsgBox Err.Description
  ShowStatus False
End Sub

Private Sub Secondary(Optional FieldName As String)

  cmdBankUpdate.Enabled = False
  If cbBank.List(0) <> "" Then
    Set ADOsecondaryRS = New ADODB.Recordset
    If FieldName = "" Then
        cbBank.Text = cbBank.List(0)
        ADOsecondaryRS.Open "SELECT * FROM [BANK Accounts]", db, adOpenKeyset, adLockOptimistic, adCmdText
    Else
        ADOsecondaryRS.Open "SELECT * FROM [BANK Accounts] WHERE [BANK ACCT Name]='" & FieldName & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
        'MsgBox ADOsecondaryRS.RecordCount
    End If
    Dim oTxtBank As TextBox
    For Each oTxtBank In Me.txtBank
        Set oTxtBank.DataSource = ADOsecondaryRS
        If ADOsecondaryRS("" & oTxtBank.DataField & "").Type = 202 Then oTxtBank.MaxLength = ADOsecondaryRS("" & oTxtBank.DataField & "").DefinedSize
    Next
    cmdBankUpdate.Enabled = True
  End If
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
  
  Me.Width = 12435
  Me.Height = 7485
  
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
  End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo FormErr
  ShowStatus True
  
      If cbBank.List(0) <> "" Then
          Dim oTxtBank As TextBox
          For Each oTxtBank In Me.txtBank
              Set oTxtBank.DataSource = Nothing
          Next
          ADOsecondaryRS.CancelUpdate
          ADOsecondaryRS.Close
          Set ADOsecondaryRS = Nothing
      End If
      
      If ADOprimaryrs.RecordCount > 0 Then
        If ADOprimaryrs.EditMode <> 0 Then
          ADOprimaryrs.CancelUpdate
        End If
      End If
      ADOprimaryrs.Close
      Set ADOprimaryrs = Nothing
      db.Close
      Set db = Nothing
      Set frm_SYS_Setup_Banks = Nothing
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
  .Update
  .Requery
  End With
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

Private Sub txtBank_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
Case 1
    keyResponse = CtrlValidate(KeyAscii, "0123456789-")
    If keyResponse = False Then
       KeyAscii = 0
    End If
Case 12, 13
    keyResponse = CtrlValidate(KeyAscii, "0123456789")
    If keyResponse = False Then
       KeyAscii = 0
    End If
Case 15, 16
    keyResponse = CtrlValidate(KeyAscii, "0123456789.")
    If keyResponse = False Then
       KeyAscii = 0
    End If
End Select
End Sub

Private Sub txtBank_LostFocus(Index As Integer)
Select Case Index
Case 15, 16
    txtBank(Index) = FormatCurr(txtBank(Index))
End Select
End Sub

Private Sub txtfields_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
Case 6, 7, 8, 9
    keyResponse = CtrlValidate(KeyAscii, "0123456789.")
    If keyResponse = True Then
    Else
       KeyAscii = 0
    End If
End Select
End Sub

Private Sub txtFields_LostFocus(Index As Integer)
Select Case Index
Case 6, 8, 9
    txtFields(Index) = Format(txtFields(Index), "00.00")
Case 7
    txtFields(Index) = FormatCurr(txtFields(Index))
End Select
End Sub
