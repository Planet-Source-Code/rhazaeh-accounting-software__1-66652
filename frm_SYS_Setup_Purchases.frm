VERSION 5.00
Begin VB.Form frm_SYS_Setup_Purchases 
   Caption         =   "Purchasing Setup"
   ClientHeight    =   4710
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   10755
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4710
   ScaleWidth      =   10755
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   10695
      TabIndex        =   52
      Top             =   0
      Width           =   10695
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "General Ledger Purchace Account"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   54
         Top             =   120
         Width           =   5175
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Purchase Particular"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   0
         Left            =   5400
         TabIndex        =   53
         Top             =   120
         Width           =   5175
      End
   End
   Begin VB.Frame frPrimary 
      Height          =   3855
      Left            =   0
      TabIndex        =   31
      Top             =   480
      Width           =   10695
      Begin VB.Frame Frame1 
         Height          =   3615
         Left            =   120
         TabIndex        =   43
         Top             =   120
         Width           =   5175
         Begin VB.CommandButton btPurchaseWriteOff 
            Height          =   285
            Left            =   4440
            Picture         =   "frm_SYS_Setup_Purchases.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   3000
            Width           =   375
         End
         Begin VB.CommandButton btGLPurchaseMisc 
            Height          =   285
            Left            =   4440
            Picture         =   "frm_SYS_Setup_Purchases.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   2640
            Width           =   375
         End
         Begin VB.CommandButton btGLPurchaseDisc 
            Height          =   285
            Left            =   4440
            Picture         =   "frm_SYS_Setup_Purchases.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   2280
            Width           =   375
         End
         Begin VB.CommandButton btGLPurchaseFreight 
            Height          =   285
            Left            =   4440
            Picture         =   "frm_SYS_Setup_Purchases.frx":091E
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   1920
            Width           =   375
         End
         Begin VB.CommandButton btGLPurchasePrePaid 
            Height          =   285
            Left            =   4440
            Picture         =   "frm_SYS_Setup_Purchases.frx":0C28
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   1560
            Width           =   375
         End
         Begin VB.CommandButton btGLPurchaseCashAcc 
            Height          =   285
            Left            =   4440
            Picture         =   "frm_SYS_Setup_Purchases.frx":0F32
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   1200
            Width           =   375
         End
         Begin VB.CommandButton btGLPurchaseInv 
            Height          =   285
            Left            =   4440
            Picture         =   "frm_SYS_Setup_Purchases.frx":123C
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   840
            Width           =   375
         End
         Begin VB.CommandButton btPurchaseAP 
            Height          =   285
            Left            =   4440
            Picture         =   "frm_SYS_Setup_Purchases.frx":1546
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox txtFields 
            DataField       =   "SYS COM Purchase PrePaid Acct"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   4
            Left            =   2880
            TabIndex        =   4
            Top             =   1560
            Width           =   1575
         End
         Begin VB.TextBox txtFields 
            DataField       =   "SYS COM Purchase Write Off Acct"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   8
            Left            =   2880
            TabIndex        =   8
            Top             =   3000
            Width           =   1575
         End
         Begin VB.TextBox txtFields 
            DataField       =   "SYS COM Purchase AP Acct"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   1
            Left            =   2880
            TabIndex        =   1
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox txtFields 
            DataField       =   "SYS COM Purchase Cash Acct"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   3
            Left            =   2880
            TabIndex        =   3
            Top             =   1200
            Width           =   1575
         End
         Begin VB.TextBox txtFields 
            DataField       =   "SYS COM Purchase Discount Acct"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   6
            Left            =   2880
            TabIndex        =   6
            Top             =   2280
            Width           =   1575
         End
         Begin VB.TextBox txtFields 
            DataField       =   "SYS COM Purchase Freight Acct"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   5
            Left            =   2880
            TabIndex        =   5
            Top             =   1920
            Width           =   1575
         End
         Begin VB.TextBox txtFields 
            DataField       =   "SYS COM Purchase Inventory Acct"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   2
            Left            =   2880
            TabIndex        =   2
            Top             =   840
            Width           =   1575
         End
         Begin VB.TextBox txtFields 
            DataField       =   "SYS COM Purchase Misc Acct"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   7
            Left            =   2880
            TabIndex        =   7
            Top             =   2640
            Width           =   1575
         End
         Begin VB.Label lblfields 
            Alignment       =   1  'Right Justify
            Caption         =   "GL Purchase - Pre Paid Account:  "
            Height          =   255
            Index           =   4
            Left            =   360
            TabIndex        =   51
            Top             =   1560
            Width           =   2535
         End
         Begin VB.Label lblfields 
            Alignment       =   1  'Right Justify
            Caption         =   "GL Purchase - Write Off Account:  "
            Height          =   255
            Index           =   8
            Left            =   360
            TabIndex        =   50
            Top             =   3000
            Width           =   2535
         End
         Begin VB.Label lblfields 
            Alignment       =   1  'Right Justify
            Caption         =   "GL Purchase - AP Account:  "
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   49
            Top             =   480
            Width           =   2535
         End
         Begin VB.Label lblfields 
            Alignment       =   1  'Right Justify
            Caption         =   "GL Purchase - Cash Account:  "
            Height          =   255
            Index           =   3
            Left            =   360
            TabIndex        =   48
            Top             =   1200
            Width           =   2535
         End
         Begin VB.Label lblfields 
            Alignment       =   1  'Right Justify
            Caption         =   "GL Purchase - Discount Account:  "
            Height          =   255
            Index           =   6
            Left            =   360
            TabIndex        =   47
            Top             =   2280
            Width           =   2535
         End
         Begin VB.Label lblfields 
            Alignment       =   1  'Right Justify
            Caption         =   "GL Purchase - Freight Account:  "
            Height          =   255
            Index           =   5
            Left            =   360
            TabIndex        =   46
            Top             =   1920
            Width           =   2535
         End
         Begin VB.Label lblfields 
            Alignment       =   1  'Right Justify
            Caption         =   "GL Purchase - Inventory Account:  "
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   45
            Top             =   840
            Width           =   2535
         End
         Begin VB.Label lblfields 
            Alignment       =   1  'Right Justify
            Caption         =   "GL Purchase - Misc Account:  "
            Height          =   255
            Index           =   7
            Left            =   360
            TabIndex        =   44
            Top             =   2640
            Width           =   2535
         End
      End
      Begin VB.Frame Frame2 
         Height          =   3615
         Left            =   5385
         TabIndex        =   32
         Top             =   120
         Width           =   5175
         Begin VB.CommandButton btcbRefresh 
            Height          =   280
            Index           =   21
            Left            =   4680
            Picture         =   "frm_SYS_Setup_Purchases.frx":1850
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Update the Ship Via"
            Top             =   720
            Width           =   375
         End
         Begin VB.CommandButton btcbRefresh 
            Height          =   280
            Index           =   20
            Left            =   2160
            Picture         =   "frm_SYS_Setup_Purchases.frx":1B5A
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Update the Ship Via"
            Top             =   720
            Width           =   375
         End
         Begin VB.ComboBox cbfields 
            DataField       =   "SYS COM Purchase Payment Methods"
            Height          =   315
            Index           =   21
            Left            =   2760
            TabIndex        =   19
            Top             =   720
            Width           =   1935
         End
         Begin VB.ComboBox cbfields 
            DataField       =   "SYS COM Purchase Payment Terms"
            Height          =   315
            Index           =   20
            Left            =   240
            TabIndex        =   17
            Top             =   720
            Width           =   1935
         End
         Begin VB.CommandButton btcbRefresh 
            Height          =   280
            Index           =   22
            Left            =   2160
            Picture         =   "frm_SYS_Setup_Purchases.frx":1E64
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Update the Ship Via"
            Top             =   1440
            Width           =   375
         End
         Begin VB.ComboBox cbfields 
            DataField       =   "SYS COM Purchase Shipping Method"
            Height          =   315
            Index           =   22
            Left            =   240
            TabIndex        =   21
            Top             =   1440
            Width           =   1935
         End
         Begin VB.TextBox txtFields 
            DataField       =   "SYS COM Purchase Period 1"
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
            Index           =   12
            Left            =   1200
            TabIndex        =   23
            Top             =   2160
            Width           =   495
         End
         Begin VB.TextBox txtFields 
            DataField       =   "SYS COM Purchase Period 2"
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
            Left            =   2880
            TabIndex        =   24
            Top             =   2160
            Width           =   495
         End
         Begin VB.TextBox txtFields 
            DataField       =   "SYS COM Purchase Period 3"
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
            Index           =   14
            Left            =   4560
            TabIndex        =   25
            Top             =   2160
            Width           =   495
         End
         Begin VB.Frame Frame3 
            Height          =   975
            Left            =   120
            TabIndex        =   33
            Top             =   2520
            Width           =   4935
            Begin VB.OptionButton optAgeInvoice 
               Caption         =   "Invoice Date"
               Height          =   255
               Index           =   1
               Left            =   720
               TabIndex        =   26
               Top             =   480
               Width           =   1575
            End
            Begin VB.OptionButton optAgeInvoice 
               Caption         =   "Due Date"
               Height          =   255
               Index           =   2
               Left            =   2760
               TabIndex        =   27
               Top             =   480
               Width           =   1575
            End
            Begin VB.Label Label2 
               Caption         =   "Age Invoices By"
               Height          =   255
               Left            =   120
               TabIndex        =   55
               Top             =   0
               Width           =   1215
            End
         End
         Begin VB.Label lblfields 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Purchase Payment Method"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   21
            Left            =   2760
            TabIndex        =   42
            Top             =   480
            Width           =   2295
         End
         Begin VB.Label lblfields 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Purchase Payment Terms"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   20
            Left            =   240
            TabIndex        =   41
            Top             =   480
            Width           =   2295
         End
         Begin VB.Label lblfields 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Purchase Shipping Method"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   22
            Left            =   240
            TabIndex        =   40
            Top             =   1200
            Width           =   2295
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Purchase Period 2"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   13
            Left            =   1920
            TabIndex        =   39
            Top             =   1905
            Width           =   1455
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Purchase Period 3"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   14
            Left            =   3600
            TabIndex        =   38
            Top             =   1905
            Width           =   1455
         End
         Begin VB.Label lblPeriod 
            Alignment       =   2  'Center
            Caption         =   "0 to"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   37
            Top             =   2205
            Width           =   975
         End
         Begin VB.Label lblPeriod 
            Alignment       =   2  'Center
            Caption         =   "31 to"
            Height          =   255
            Index           =   1
            Left            =   1920
            TabIndex        =   36
            Top             =   2205
            Width           =   975
         End
         Begin VB.Label lblPeriod 
            Alignment       =   2  'Center
            Caption         =   "61 to"
            Height          =   255
            Index           =   2
            Left            =   3600
            TabIndex        =   35
            Top             =   2205
            Width           =   975
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Purchase Period 1"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   12
            Left            =   240
            TabIndex        =   34
            Top             =   1905
            Width           =   1455
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
      ScaleWidth      =   10755
      TabIndex        =   0
      Top             =   4410
      Width           =   10755
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   300
         Left            =   2160
         TabIndex        =   30
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   300
         Left            =   1080
         TabIndex        =   29
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   300
         Left            =   0
         TabIndex        =   28
         Top             =   0
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frm_SYS_Setup_Purchases"
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

Private Sub btcbRefresh_Click(Index As Integer)
    Dim tmp As String
    tmp = cbfields(Index).Text
    loadCombo Index
    cbfields(Index) = tmp
End Sub

'Private Sub btGLPurchase_Click()
'    Dim No As Integer
'    Dim sql As String
'    Dim ghead As String
'    Dim fhead As String
    
'    No = 16
'    sql = "select [GL COA Account No], [GL COA Account Name]" & _
'        "from [GL Chart of Accounts]"
'    ghead = "Account Description"
'    fhead = "Account No//Description"
    
'    AllLookup.GetWhichTable No, sql, ghead, fhead,db
'    'AllLookup.Show vbModal
'    txtFields(0).SetFocus
'End Sub

Private Sub btGLPurchaseCashAcc_Click()
    Dim No As Integer
    Dim sql As String
    Dim ghead As String
    Dim fhead As String
    
    No = 19
    sql = "select [GL COA Account No], [GL COA Account Name]" & _
        "from [GL Chart of Accounts]"
    ghead = "Account Description"
    fhead = "Account No//Description"
    
    AllLookup.GetWhichTable No, sql, ghead, fhead, db
    'AllLookup.Show vbModal
    txtfields(3).SetFocus
End Sub

Private Sub btGLPurchaseDisc_Click()
    Dim No As Integer
    Dim sql As String
    Dim ghead As String
    Dim fhead As String
    
    No = 22
    sql = "select [GL COA Account No], [GL COA Account Name]" & _
        "from [GL Chart of Accounts]"
    ghead = "Account Description"
    fhead = "Account No//Description"
    
    AllLookup.GetWhichTable No, sql, ghead, fhead, db
    'AllLookup.Show vbModal
    txtfields(6).SetFocus
End Sub

Private Sub btGLPurchaseFreight_Click()
    Dim No As Integer
    Dim sql As String
    Dim ghead As String
    Dim fhead As String
    
    No = 21
    sql = "select [GL COA Account No], [GL COA Account Name]" & _
        "from [GL Chart of Accounts]"
    ghead = "Account Description"
    fhead = "Account No//Description"
    
    AllLookup.GetWhichTable No, sql, ghead, fhead, db
    'AllLookup.Show vbModal
    txtfields(5).SetFocus
End Sub

Private Sub btGLPurchaseInv_Click()
    Dim No As Integer
    Dim sql As String
    Dim ghead As String
    Dim fhead As String
    
    No = 18
    sql = "select [GL COA Account No], [GL COA Account Name]" & _
        "from [GL Chart of Accounts]"
    ghead = "Account Description"
    fhead = "Account No//Description"
    
    AllLookup.GetWhichTable No, sql, ghead, fhead, db
    'AllLookup.Show vbModal
    txtfields(2).SetFocus
End Sub

Private Sub btGLPurchaseMisc_Click()
    Dim No As Integer
    Dim sql As String
    Dim ghead As String
    Dim fhead As String
    
    No = 23
    sql = "select [GL COA Account No], [GL COA Account Name]" & _
        "from [GL Chart of Accounts]"
    ghead = "Account Description"
    fhead = "Account No//Description"
    
    AllLookup.GetWhichTable No, sql, ghead, fhead, db
    'AllLookup.Show vbModal
    txtfields(7).SetFocus
End Sub

Private Sub btGLPurchasePrePaid_Click()
    Dim No As Integer
    Dim sql As String
    Dim ghead As String
    Dim fhead As String
    
    No = 20
    sql = "select [GL COA Account No], [GL COA Account Name]" & _
        "from [GL Chart of Accounts]"
    ghead = "Account Description"
    fhead = "Account No//Description"
    
    AllLookup.GetWhichTable No, sql, ghead, fhead, db
    'AllLookup.Show vbModal
    txtfields(4).SetFocus
End Sub

Private Sub btPurchaseAP_Click()
    Dim No As Integer
    Dim sql As String
    Dim ghead As String
    Dim fhead As String
    
    No = 17
    sql = "select [GL COA Account No], [GL COA Account Name]" & _
        "from [GL Chart of Accounts]"
    ghead = "Account Description"
    fhead = "Account No//Description"
    
    AllLookup.GetWhichTable No, sql, ghead, fhead, db
    'AllLookup.Show vbModal
    txtfields(1).SetFocus
End Sub

Private Sub btPurchaseWriteOff_Click()
    Dim No As Integer
    Dim sql As String
    Dim ghead As String
    Dim fhead As String
    
    No = 24
    sql = "select [GL COA Account No], [GL COA Account Name]" & _
        "from [GL Chart of Accounts]"
    ghead = "Account Description"
    fhead = "Account No//Description"
    
    AllLookup.GetWhichTable No, sql, ghead, fhead, db
    'AllLookup.Show vbModal
    txtfields(8).SetFocus
End Sub

Private Sub cbfields_LostFocus(Index As Integer)
Select Case Index
Case 1
   CheckCombo cbfields(Index), "[LIST PAY Description]", "[LIST Payment Terms]", db, True
Case 2
   CheckCombo cbfields(Index), "[LIST PAY Method]", "[LIST Payment Methods]", db, True
Case 3
   CheckCombo cbfields(Index), "[LIST SHIP Method]", "[LIST Shipping Methods]", db, True
End Select
End Sub

Private Sub Form_Load()
ShowStatus True
'On Error GoTo FormErr
NewLoad = True
  Set db = New ADODB.Connection
  db.CursorLocation = adUseClient
  db.Open gblADOProvider

  Set ADOprimaryrs = New ADODB.Recordset
  ADOprimaryrs.Open "select [SYS COM Purchase Number],[SYS COM Purchase Payment Methods]," & _
    "[SYS COM Purchase Payment Terms],[SYS COM Purchase Period 1]," & _
    "[SYS COM Purchase Period 2],[SYS COM Purchase Period 3]," & _
    "[SYS COM Purchase PrePaid Acct],[SYS COM Purchase Shipping Method]," & _
    "[SYS COM Purchase Write Off Acct],[SYS COM Purchase Age Invoices By]," & _
    "[SYS COM Purchase AP Acct],[SYS COM Purchase Cash Acct]," & _
    "[SYS COM Purchase Discount Acct],[SYS COM Purchase Freight Acct]," & _
    "[SYS COM Purchase Inventory Acct],[SYS COM Purchase Misc Acct] " & _
    "from [SYS Company]", db, adOpenStatic, adLockOptimistic

  Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtfields
    Set oText.DataSource = ADOprimaryrs
  Next
  
  Dim oCombo As ComboBox
  'Bind Datacombos to the data provider
  For Each oCombo In Me.cbfields
    Set oCombo.DataSource = ADOprimaryrs
  Next
    loadCombo
  
  If CheckNewDB(ADOprimaryrs, "Purchases") = True Then
    ADOprimaryrs.AddNew
  End If
    
  If ADOprimaryrs![SYS COM Purchase Age Invoices By] = 1 Then
    optAgeInvoice(1).Value = True
  ElseIf ADOprimaryrs![SYS COM Purchase Age Invoices By] = 2 Then
    optAgeInvoice(2).Value = True
  End If
  If txtfields(12) = "0" Then
    txtfields(12) = "30"
    txtfields(13) = "60"
    txtfields(14) = "90"
  End If
  GetTextColor Me
  
  mbDataChanged = False
NewLoad = False
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
  
  Me.Width = 10875
  Me.Height = 5115
  
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

  For Each oText In Me.txtfields
    If oText.Text = "" Then
        Response = MsgBox("All data must be filled" & vbCr & "Are you sure want to quit and leave all of this important data empty?", vbYesNo, "Error")
        If Response = vbNo Then
            Cancel = 1
            Exit Sub
        Else
            Exit For
        End If
    End If
  Next
  ShowStatus True
      EndLoad db, ADOprimaryrs, "Purchasing Preferences"
      If ADOprimaryrs.RecordCount > 0 Then
        If ADOprimaryrs.EditMode <> 0 Then
          ADOprimaryrs.CancelUpdate
        End If
      End If
      ADOprimaryrs.Close
      Set ADOprimaryrs = Nothing
      db.Close
      Set db = Nothing
  Set frm_SYS_Setup_Purchases = Nothing
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
  '.Requery
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

Private Sub loadCombo(Optional Index As Integer)
    Index = IIf(Index > 0, Index, 0)
    Select Case Index
    Case 0
        ComboInit cbfields(20), lblfields(20), "select [LIST PAY Description] " & _
            "from [LIST Payment Terms]"
        ComboInit cbfields(21), lblfields(21), "select [LIST PAY Method] " & _
            "from [LIST Payment Methods]"
        ComboInit cbfields(22), lblfields(22), "select [LIST SHIP Method] " & _
            "from [LIST Shipping Methods]"
    Case 1
        ComboInit cbfields(20), lblfields(20), "select [LIST PAY Description] " & _
            "from [LIST Payment Terms]"
    Case 2
        ComboInit cbfields(21), lblfields(21), "select [LIST PAY Method] " & _
            "from [LIST Payment Methods]"
    Case 3
        ComboInit cbfields(22), lblfields(22), "select [LIST SHIP Method] " & _
            "from [LIST Shipping Methods]"
    End Select
End Sub

Private Sub optAgeInvoice_Click(Index As Integer)
    If NewLoad = False Then ADOprimaryrs![SYS COM Purchase Age Invoices By] = Index
End Sub

Private Sub txtfields_KeyPress(Index As Integer, KeyAscii As Integer)
    keyResponse = CtrlValidate(KeyAscii, "0123456789")
    If keyResponse = True Then
    Else
       KeyAscii = 0
    End If
End Sub

Private Sub txtFields_LostFocus(Index As Integer)
If Trim(txtfields(Index).Text) = "" Then Exit Sub

Select Case Index
Case 1, 2, 3, 4, 5, 6, 7, 8
    If IsNumeric(txtfields(Index).Text) And txtfields(Index).Text <> "" Then
        CheckDocument "SELECT [GL COA Account No] FROM [GL Chart Of Accounts] WHERE [GL COA Account No]='" & txtfields(Index).Text & "'", db, False, txtfields(Index), "COA"
    Else
        If txtfields(Index).Text <> "" Then MsgBox "Only numeric character is accepted", vbInformation, "Information"
        txtfields(Index) = ""
    End If
Case 12
    lblPeriod(1) = txtfields(Index) & " to"
Case 13
    lblPeriod(2) = txtfields(Index) & " to"
End Select
End Sub
