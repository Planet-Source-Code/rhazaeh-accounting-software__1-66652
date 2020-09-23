VERSION 5.00
Begin VB.Form frm_SYS_Setup_Sales 
   Caption         =   "Sales Setup"
   ClientHeight    =   5685
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   10515
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5685
   ScaleWidth      =   10515
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   10455
      TabIndex        =   36
      Top             =   0
      Width           =   10455
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "General Ledger Sales Account"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   38
         Top             =   120
         Width           =   5055
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sales Particular"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   0
         Left            =   5280
         TabIndex        =   37
         Top             =   120
         Width           =   5055
      End
   End
   Begin VB.Frame frPrimary 
      Height          =   4815
      Left            =   0
      TabIndex        =   23
      Top             =   480
      Width           =   10455
      Begin VB.CheckBox Check1 
         Caption         =   "Charge Interest On Overdue Acct."
         DataField       =   "SYS COM Finance Charges YN"
         Height          =   195
         Left            =   5400
         TabIndex        =   59
         Top             =   240
         Width           =   2775
      End
      Begin VB.Frame Frame2 
         Height          =   2775
         Left            =   5280
         TabIndex        =   58
         Top             =   240
         Width           =   3135
         Begin VB.TextBox txtFields 
            DataField       =   "SYS COM Minimum Finance Charge"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   14
            Left            =   1560
            TabIndex        =   67
            Top             =   2160
            Width           =   1335
         End
         Begin VB.TextBox txtFields 
            DataField       =   "SYS COM Minimum Balance"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   13
            Left            =   1560
            TabIndex        =   66
            Top             =   1680
            Width           =   1335
         End
         Begin VB.TextBox txtFields 
            DataField       =   "SYS COM Monthly Charge"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   12
            Left            =   1560
            TabIndex        =   65
            Top             =   1200
            Width           =   975
         End
         Begin VB.TextBox txtFields 
            DataField       =   "SYS COM Annual Charge"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   11
            Left            =   1560
            TabIndex        =   64
            Top             =   720
            Width           =   975
         End
         Begin VB.Label lblLabels 
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
            Height          =   255
            Index           =   6
            Left            =   2640
            TabIndex        =   69
            Top             =   1200
            Width           =   255
         End
         Begin VB.Label lblLabels 
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
            Height          =   255
            Index           =   5
            Left            =   2640
            TabIndex        =   68
            Top             =   720
            Width           =   255
         End
         Begin VB.Label lblLabels 
            Caption         =   "Minimum Charges"
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   63
            Top             =   2160
            Width           =   1335
         End
         Begin VB.Label lblLabels 
            Caption         =   "Minimum Balance"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   62
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label lblLabels 
            Caption         =   "Monthly Charges"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   61
            Top             =   1200
            Width           =   1335
         End
         Begin VB.Label lblLabels 
            Caption         =   "Annual Charges"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   60
            Top             =   720
            Width           =   1335
         End
      End
      Begin VB.Frame Frame5 
         Height          =   2775
         Left            =   8520
         TabIndex        =   54
         Top             =   240
         Width           =   1815
         Begin VB.OptionButton optSalesAccDef 
            Caption         =   "Customer"
            Height          =   255
            Index           =   0
            Left            =   600
            TabIndex        =   56
            Top             =   720
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optSalesAccDef 
            Caption         =   "Item"
            Height          =   255
            Index           =   1
            Left            =   600
            TabIndex        =   55
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label frLabel 
            Caption         =   "Default Sales Acct."
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   57
            Top             =   0
            Width           =   1455
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1575
         Left            =   5280
         TabIndex        =   39
         Top             =   3120
         Width           =   5055
         Begin VB.OptionButton optAgeInv 
            Caption         =   "Due Date"
            Height          =   255
            Index           =   1
            Left            =   3720
            TabIndex        =   44
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton optAgeInv 
            Caption         =   "Invoice Date"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   43
            Top             =   360
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.TextBox txtFields 
            DataField       =   "SYS COM Sales Period 1"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   15
            Left            =   720
            TabIndex        =   42
            Top             =   1110
            Width           =   495
         End
         Begin VB.TextBox txtFields 
            DataField       =   "SYS COM Sales Period 2"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   16
            Left            =   1920
            TabIndex        =   41
            Top             =   1110
            Width           =   495
         End
         Begin VB.TextBox txtFields 
            DataField       =   "SYS COM Sales Period 3"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   17
            Left            =   3240
            TabIndex        =   40
            Top             =   1110
            Width           =   495
         End
         Begin VB.Label frLabel 
            Caption         =   "Age Invoices By"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   53
            Top             =   0
            Width           =   1215
         End
         Begin VB.Label lblfields 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000001&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Aging Period 1"
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   52
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label lblPeriod 
            Alignment       =   2  'Center
            Caption         =   "0 to"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   51
            Top             =   1140
            Width           =   615
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000001&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Aging Period 2"
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   12
            Left            =   1320
            TabIndex        =   50
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label lblPeriod 
            Alignment       =   2  'Center
            Caption         =   "30 to"
            Height          =   255
            Index           =   1
            Left            =   1320
            TabIndex        =   49
            Top             =   1140
            Width           =   615
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000001&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Aging Period 3"
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   13
            Left            =   2640
            TabIndex        =   48
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label lblPeriod 
            Alignment       =   2  'Center
            Caption         =   "60 to"
            Height          =   255
            Index           =   2
            Left            =   2640
            TabIndex        =   47
            Top             =   1140
            Width           =   615
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000001&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Aging Period 4"
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   0
            Left            =   3840
            TabIndex        =   46
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label lblPeriod 
            Alignment       =   2  'Center
            Caption         =   "Over 90 days"
            Height          =   255
            Index           =   3
            Left            =   3840
            TabIndex        =   45
            Top             =   1140
            Width           =   975
         End
      End
      Begin VB.Frame Frame1 
         Height          =   4575
         Left            =   120
         TabIndex        =   25
         Top             =   120
         Width           =   5055
         Begin VB.TextBox txtFields 
            DataField       =   "SYS COM Finance Charge Acct"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   0
            Left            =   2520
            TabIndex        =   71
            Top             =   3960
            Width           =   1455
         End
         Begin VB.CommandButton Command1 
            Height          =   285
            Left            =   3960
            Picture         =   "frm_SYS_Setup_Sales.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   70
            Top             =   3960
            Width           =   375
         End
         Begin VB.CommandButton btGLSalesWriteOff 
            Height          =   285
            Left            =   3960
            Picture         =   "frm_SYS_Setup_Sales.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   3600
            Width           =   375
         End
         Begin VB.CommandButton btGLSalesReturn 
            Height          =   285
            Left            =   3960
            Picture         =   "frm_SYS_Setup_Sales.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   3240
            Width           =   375
         End
         Begin VB.CommandButton btGLSalesSales 
            Height          =   285
            Left            =   3960
            Picture         =   "frm_SYS_Setup_Sales.frx":091E
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   2880
            Width           =   375
         End
         Begin VB.CommandButton btGLSalesMisc 
            Height          =   285
            Left            =   3960
            Picture         =   "frm_SYS_Setup_Sales.frx":0C28
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   2520
            Width           =   375
         End
         Begin VB.CommandButton btGLSalesInv 
            Height          =   285
            Left            =   3960
            Picture         =   "frm_SYS_Setup_Sales.frx":0F32
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   2160
            Width           =   375
         End
         Begin VB.CommandButton btGLSalesFreight 
            Height          =   285
            Left            =   3960
            Picture         =   "frm_SYS_Setup_Sales.frx":123C
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   1800
            Width           =   375
         End
         Begin VB.CommandButton btGLSalesDisc 
            Height          =   285
            Left            =   3960
            Picture         =   "frm_SYS_Setup_Sales.frx":1546
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   1440
            Width           =   375
         End
         Begin VB.CommandButton btGLSalesCogs 
            Height          =   285
            Left            =   3960
            Picture         =   "frm_SYS_Setup_Sales.frx":1850
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   1080
            Width           =   375
         End
         Begin VB.CommandButton btGLSalesCash 
            Height          =   285
            Left            =   3960
            Picture         =   "frm_SYS_Setup_Sales.frx":1B5A
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   720
            Width           =   375
         End
         Begin VB.CommandButton btGLSalesAR 
            Height          =   285
            Left            =   3960
            Picture         =   "frm_SYS_Setup_Sales.frx":1E64
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   360
            Width           =   375
         End
         Begin VB.TextBox txtFields 
            DataField       =   "SYS COM Sales AR Acct"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   1
            Left            =   2520
            TabIndex        =   0
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox txtFields 
            DataField       =   "SYS COM Sales Cash Acct"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   2
            Left            =   2520
            TabIndex        =   1
            Top             =   720
            Width           =   1455
         End
         Begin VB.TextBox txtFields 
            DataField       =   "SYS COM Sales COGS Acct"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   3
            Left            =   2520
            TabIndex        =   2
            Top             =   1080
            Width           =   1455
         End
         Begin VB.TextBox txtFields 
            DataField       =   "SYS COM Sales Discount Acct"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   4
            Left            =   2520
            TabIndex        =   3
            Top             =   1440
            Width           =   1455
         End
         Begin VB.TextBox txtFields 
            DataField       =   "SYS COM Sales Freight Acct"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   5
            Left            =   2520
            TabIndex        =   4
            Top             =   1800
            Width           =   1455
         End
         Begin VB.TextBox txtFields 
            DataField       =   "SYS COM Sales Inventory Acct"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   6
            Left            =   2520
            TabIndex        =   5
            Top             =   2160
            Width           =   1455
         End
         Begin VB.TextBox txtFields 
            DataField       =   "SYS COM Sales Misc Acct"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   7
            Left            =   2520
            TabIndex        =   6
            Top             =   2520
            Width           =   1455
         End
         Begin VB.TextBox txtFields 
            DataField       =   "SYS COM Sales Return Acct"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   9
            Left            =   2520
            TabIndex        =   8
            Top             =   3240
            Width           =   1455
         End
         Begin VB.TextBox txtFields 
            DataField       =   "SYS COM Sales Sales Acct"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   8
            Left            =   2520
            TabIndex        =   7
            Top             =   2880
            Width           =   1455
         End
         Begin VB.TextBox txtFields 
            DataField       =   "SYS COM Sales Write Off Acct"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   10
            Left            =   2520
            TabIndex        =   9
            Top             =   3600
            Width           =   1455
         End
         Begin VB.Label lblfields 
            Alignment       =   1  'Right Justify
            Caption         =   "GL Sales - Finance Charge  "
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   72
            Top             =   3960
            Width           =   2175
         End
         Begin VB.Label lblfields 
            Alignment       =   1  'Right Justify
            Caption         =   "GL Sales - AR Account  "
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   35
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label lblfields 
            Alignment       =   1  'Right Justify
            Caption         =   "GL Sales - Cash Account  "
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   34
            Top             =   720
            Width           =   2175
         End
         Begin VB.Label lblfields 
            Alignment       =   1  'Right Justify
            Caption         =   "GL Sales - Cogs Account  "
            Height          =   255
            Index           =   3
            Left            =   360
            TabIndex        =   33
            Top             =   1080
            Width           =   2175
         End
         Begin VB.Label lblfields 
            Alignment       =   1  'Right Justify
            Caption         =   "GL Sales - Discount Account  "
            Height          =   255
            Index           =   4
            Left            =   360
            TabIndex        =   32
            Top             =   1440
            Width           =   2175
         End
         Begin VB.Label lblfields 
            Alignment       =   1  'Right Justify
            Caption         =   "Gl Sales - Freight Account  "
            Height          =   255
            Index           =   5
            Left            =   360
            TabIndex        =   31
            Top             =   1800
            Width           =   2175
         End
         Begin VB.Label lblfields 
            Alignment       =   1  'Right Justify
            Caption         =   "GL Sales - Inventory Account  "
            Height          =   255
            Index           =   6
            Left            =   360
            TabIndex        =   30
            Top             =   2160
            Width           =   2175
         End
         Begin VB.Label lblfields 
            Alignment       =   1  'Right Justify
            Caption         =   "GL Sales - Non Stock Sales"
            Height          =   255
            Index           =   7
            Left            =   360
            TabIndex        =   29
            Top             =   2520
            Width           =   2175
         End
         Begin VB.Label lblfields 
            Alignment       =   1  'Right Justify
            Caption         =   "GL Sales - Returns Account  "
            Height          =   255
            Index           =   9
            Left            =   240
            TabIndex        =   28
            Top             =   3240
            Width           =   2295
         End
         Begin VB.Label lblfields 
            Alignment       =   1  'Right Justify
            Caption         =   "GL Sales - Sales Account  "
            Height          =   255
            Index           =   8
            Left            =   360
            TabIndex        =   27
            Top             =   2880
            Width           =   2175
         End
         Begin VB.Label lblfields 
            Alignment       =   1  'Right Justify
            Caption         =   "GL Sales - Write Off Account  "
            Height          =   255
            Index           =   10
            Left            =   360
            TabIndex        =   26
            Top             =   3600
            Width           =   2175
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
      ScaleWidth      =   10515
      TabIndex        =   20
      Top             =   5385
      Width           =   10515
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   300
         Left            =   2160
         TabIndex        =   24
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   300
         Left            =   1080
         TabIndex        =   22
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   300
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frm_SYS_Setup_Sales"
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

Private Sub btGLSalesAR_Click()
    Dim No As Integer
    Dim sql As String
    Dim ghead As String
    Dim fhead As String
    
    No = 26
    sql = "select [GL COA Account No], [GL COA Account Name]" & _
        "from [GL Chart of Accounts]"
    ghead = "Account Description"
    fhead = "Account No//Description"
    
    AllLookup.GetWhichTable No, sql, ghead, fhead, db
    'AllLookup.Show vbModal
    txtfields(1).SetFocus
End Sub

Private Sub btGLSalesCash_Click()
    Dim No As Integer
    Dim sql As String
    Dim ghead As String
    Dim fhead As String
    
    No = 27
    sql = "select [GL COA Account No], [GL COA Account Name]" & _
        "from [GL Chart of Accounts]"
    ghead = "Account Description"
    fhead = "Account No//Description"
    
    AllLookup.GetWhichTable No, sql, ghead, fhead, db
    'AllLookup.Show vbModal
    txtfields(2).SetFocus
End Sub

Private Sub btGLSalesCogs_Click()
    Dim No As Integer
    Dim sql As String
    Dim ghead As String
    Dim fhead As String
    
    No = 28
    sql = "select [GL COA Account No], [GL COA Account Name]" & _
        "from [GL Chart of Accounts]"
    ghead = "Account Description"
    fhead = "Account No//Description"
    
    AllLookup.GetWhichTable No, sql, ghead, fhead, db
    'AllLookup.Show vbModal
    txtfields(3).SetFocus
End Sub

Private Sub btGLSalesDef_Click()
    Dim No As Integer
    Dim sql As String
    Dim ghead As String
    Dim fhead As String
    
    No = 25
    sql = "select [GL COA Account No], [GL COA Account Name]" & _
        "from [GL Chart of Accounts]"
    ghead = "Account Description"
    fhead = "Account No//Description"
    
    AllLookup.GetWhichTable No, sql, ghead, fhead, db
    'AllLookup.Show vbModal
    txtfields(0).SetFocus
End Sub

Private Sub btGLSalesDisc_Click()
    Dim No As Integer
    Dim sql As String
    Dim ghead As String
    Dim fhead As String
    
    No = 29
    sql = "select [GL COA Account No], [GL COA Account Name]" & _
        "from [GL Chart of Accounts]"
    ghead = "Account Description"
    fhead = "Account No//Description"
    
    AllLookup.GetWhichTable No, sql, ghead, fhead, db
    'AllLookup.Show vbModal
    txtfields(4).SetFocus
End Sub

Private Sub btGLSalesFreight_Click()
    Dim No As Integer
    Dim sql As String
    Dim ghead As String
    Dim fhead As String
    
    No = 30
    sql = "select [GL COA Account No], [GL COA Account Name]" & _
        "from [GL Chart of Accounts]"
    ghead = "Account Description"
    fhead = "Account No//Description"
    
    AllLookup.GetWhichTable No, sql, ghead, fhead, db
    'AllLookup.Show vbModal
    txtfields(5).SetFocus
End Sub

Private Sub btGLSalesInv_Click()
    Dim No As Integer
    Dim sql As String
    Dim ghead As String
    Dim fhead As String
    
    No = 31
    sql = "select [GL COA Account No], [GL COA Account Name]" & _
        "from [GL Chart of Accounts]"
    ghead = "Account Description"
    fhead = "Account No//Description"
    
    AllLookup.GetWhichTable No, sql, ghead, fhead, db
    'AllLookup.Show vbModal
    txtfields(6).SetFocus
End Sub

Private Sub btGLSalesMisc_Click()
    Dim No As Integer
    Dim sql As String
    Dim ghead As String
    Dim fhead As String
    
    No = 32
    sql = "select [GL COA Account No], [GL COA Account Name]" & _
        "from [GL Chart of Accounts]"
    ghead = "Account Description"
    fhead = "Account No//Description"
    
    AllLookup.GetWhichTable No, sql, ghead, fhead, db
    'AllLookup.Show vbModal
    txtfields(7).SetFocus
End Sub

Private Sub btGLSalesReturn_Click()
    Dim No As Integer
    Dim sql As String
    Dim ghead As String
    Dim fhead As String
    
    No = 34
    sql = "select [GL COA Account No], [GL COA Account Name]" & _
        "from [GL Chart of Accounts]"
    ghead = "Account Description"
    fhead = "Account No//Description"
    
    AllLookup.GetWhichTable No, sql, ghead, fhead, db
    'AllLookup.Show vbModal
    txtfields(9).SetFocus
End Sub

Private Sub btGLSalesSales_Click()
    Dim No As Integer
    Dim sql As String
    Dim ghead As String
    Dim fhead As String
    
    No = 33
    sql = "select [GL COA Account No], [GL COA Account Name]" & _
        "from [GL Chart of Accounts]"
    ghead = "Account Description"
    fhead = "Account No//Description"
    
    AllLookup.GetWhichTable No, sql, ghead, fhead, db
    'AllLookup.Show vbModal
    txtfields(8).SetFocus
End Sub

Private Sub btGLSalesTax_Click()
    Dim No As Integer
    Dim sql As String
    Dim ghead As String
    Dim fhead As String
    
    No = 34
    sql = "select [GL COA Account No], [GL COA Account Name]" & _
        "from [GL Chart of Accounts]"
    ghead = "Account Description"
    fhead = "Account No//Description"
    
    AllLookup.GetWhichTable No, sql, ghead, fhead, db
    'AllLookup.Show vbModal
    txtfields(9).SetFocus
End Sub

Private Sub btGLSalesWriteOff_Click()
    Dim No As Integer
    Dim sql As String
    Dim ghead As String
    Dim fhead As String
    
    No = 35
    sql = "select [GL COA Account No], [GL COA Account Name]" & _
        "from [GL Chart of Accounts]"
    ghead = "Account Description"
    fhead = "Account No//Description"
    
    AllLookup.GetWhichTable No, sql, ghead, fhead, db
    'AllLookup.Show vbModal
    txtfields(10).SetFocus
End Sub

Private Sub Check1_Click()
    If Check1.Value = 1 Then
        Frame2.Enabled = True
    Else
        Frame2.Enabled = False
    End If
    GetTextColor Me
End Sub

Private Sub Command1_Click()
    Dim No As Integer
    Dim sql As String
    Dim ghead As String
    Dim fhead As String
    
    No = 36
    sql = "select [GL COA Account No], [GL COA Account Name]" & _
        "from [GL Chart of Accounts]"
    ghead = "Account Description"
    fhead = "Account No//Description"
    
    AllLookup.GetWhichTable No, sql, ghead, fhead, db
    'AllLookup.Show vbModal
    txtfields(0).SetFocus

End Sub

Private Sub Form_Load()
ShowStatus True
'On Error GoTo FormErr
  Set db = New ADODB.Connection
  db.CursorLocation = adUseClient
  db.Open gblADOProvider

  Set ADOprimaryrs = New ADODB.Recordset
  ADOprimaryrs.Open "SELECT [SYS COM Annual Charge],[SYS COM Monthly Charge],[SYS COM Minimum Balance]," & _
  "[SYS COM Sales Acct Default],[SYS COM Sales Age Invoices By],[SYS COM Minimum Finance Charge]," & _
  "[SYS COM Sales AR Acct],[SYS COM Sales Cash Acct],[SYS COM Sales COGS Acct],[SYS COM Sales Discount Acct]," & _
  "[SYS COM Sales Freight Acct],[SYS COM Sales Inventory Acct],[SYS COM Sales Misc Acct]," & _
  "[SYS COM Sales Period 1],[SYS COM Sales Period 2],[SYS COM Sales Period 3]," & _
  "[SYS COM Sales Return Acct],[SYS COM Sales Sales Acct],[SYS COM Sales Sales Tax]," & _
  "[SYS COM Sales Write Off Acct],[SYS COM Finance Charges YN],[SYS COM Finance Charge Acct] " & _
  "FROM [SYS Company]", db, adOpenStatic, adLockOptimistic

  Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtfields
    Set oText.DataSource = ADOprimaryrs
  Next
    
    'Dim oCombo As ComboBox
    'Bind the comboboxes to the data provider
    'For Each oCombo In Me.cbfields
    Set Check1.DataSource = ADOprimaryrs
    'Next
    If ADOprimaryrs.RecordCount > 0 Then
        'Bind the option buttons to the data provider
      If ADOprimaryrs.Fields("SYS COM Sales Acct Default") > 0 Then _
         optSalesAccDef(ADOprimaryrs.Fields("SYS COM Sales Acct Default") - 1).Value = True
      If ADOprimaryrs.Fields("SYS COM Sales Age Invoices By") > 0 Then _
         optAgeInv(ADOprimaryrs.Fields("SYS COM Sales Age Invoices By") - 1).Value = True
    End If
  
  If CheckNewDB(ADOprimaryrs, "Sales") = True Then
    ADOprimaryrs.AddNew
  End If
  If txtfields(15) = "0" Then
    txtfields(15) = "30"
    txtfields(16) = "60"
    txtfields(17) = "90"
  End If
  
  GetTextColor Me
ShowStatus False
  mbDataChanged = False
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
  
  Me.Width = 10635
  Me.Height = 6090
  
SkipResize:
  frPrimary.Left = (Me.ScaleWidth - frPrimary.Width) / 2
  Picture1.Left = frPrimary.Left
  'lblTop.Width = frPrimary.Width
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
Dim oText As TextBox
Dim Response As Integer
'Sales Preferences
  For Each oText In Me.txtfields
    If oText.Text = "" Then
        Response = MsgBox("All data must be filled." & vbCr & "Are you sure want to quit and leave all of this important data empty?", vbYesNo, "Error")
        If Response = vbNo Then
            Cancel = 1
            Exit Sub
        Else
            Exit For
        End If
    End If
  Next
  ShowStatus True
      EndLoad db, ADOprimaryrs, "Sales Preferences"
      If ADOprimaryrs.RecordCount > 0 Then
        If ADOprimaryrs.EditMode <> 0 Then
          ADOprimaryrs.CancelUpdate
        End If
      End If
      ADOprimaryrs.Close
      Set ADOprimaryrs = Nothing
      db.Close
      Set db = Nothing
  Set frm_SYS_Setup_Sales = Nothing
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

Private Sub optAgeInv_Click(Index As Integer)
    With ADOprimaryrs
        .Fields("SYS COM Sales Age Invoices By") = Index + 1
        .Update
    End With
End Sub

Private Sub optSalesAccDef_Click(Index As Integer)
    With ADOprimaryrs
        .Fields("SYS COM Sales Acct Default") = Index + 1
        .Update
    End With
End Sub

Private Sub txtFields_Change(Index As Integer)
Select Case Index
Case 15
    lblPeriod(1) = txtfields(Index) & " to"
Case 16
    lblPeriod(2) = txtfields(Index) & " to"
Case 17
    lblPeriod(3) = "Over " & txtfields(Index)
End Select
End Sub

Private Sub txtfields_GotFocus(Index As Integer)
    TxtGotFocus txtfields(Index)
End Sub

Private Sub txtfields_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
Case 11, 12, 13, 14
    keyResponse = CtrlValidate(KeyAscii, "0123456789.")
    If keyResponse = True Then
    Else
       KeyAscii = 0
    End If
Case Else
    keyResponse = CtrlValidate(KeyAscii, "0123456789")
    If keyResponse = True Then
    Else
       KeyAscii = 0
    End If
End Select
End Sub

Private Sub txtFields_LostFocus(Index As Integer)
Select Case Index
Case 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10
    If Trim(txtfields(Index)) = "" Then
        txtfields(Index) = " "
        Exit Sub
    End If
    If IsNumeric(txtfields(Index).Text) Then
        CheckDocument "SELECT [GL COA Account No] FROM [GL Chart Of Accounts] WHERE [GL COA Account No]='" & txtfields(Index).Text & "'", db, False, txtfields(Index), "COA"
    Else
        MsgBox "Only numeric character is accepted", vbInformation, "Information"
        txtfields(Index) = " "
    End If
Case 13, 14
    txtfields(Index) = FormatCurr(txtfields(Index))
Case 11, 12
End Select
End Sub
