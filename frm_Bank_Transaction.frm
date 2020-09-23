VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_Bank_Transaction 
   Caption         =   "Bank Transaction"
   ClientHeight    =   7995
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8205
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7995
   ScaleWidth      =   8205
   Begin VB.Frame frPrimary 
      Height          =   6855
      Left            =   0
      TabIndex        =   13
      Top             =   480
      Width           =   8175
      Begin VB.PictureBox picPosted 
         BorderStyle     =   0  'None
         Height          =   525
         Left            =   5160
         Picture         =   "frm_Bank_Transaction.frx":0000
         ScaleHeight     =   525
         ScaleWidth      =   2955
         TabIndex        =   52
         Top             =   120
         Width           =   2955
      End
      Begin VB.PictureBox picMinor 
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   360
         ScaleHeight     =   975
         ScaleWidth      =   7695
         TabIndex        =   42
         Top             =   5760
         Width           =   7695
         Begin VB.TextBox txtFields 
            DataField       =   "BANK TRANS Notes"
            Height          =   555
            Index           =   14
            Left            =   1560
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   49
            Top             =   0
            Width           =   3375
         End
         Begin VB.CommandButton UnselectAll 
            Caption         =   "Unselect All"
            Height          =   375
            Left            =   6120
            TabIndex        =   48
            Top             =   0
            Width           =   975
         End
         Begin VB.CommandButton SelectAll 
            Caption         =   "Select All"
            Height          =   375
            Left            =   6120
            TabIndex        =   47
            Top             =   400
            Width           =   975
         End
         Begin VB.CommandButton cmdPost 
            Caption         =   "&Post"
            Height          =   780
            Left            =   5040
            Picture         =   "frm_Bank_Transaction.frx":0F0B
            Style           =   1  'Graphical
            TabIndex        =   46
            Top             =   0
            Width           =   855
         End
         Begin VB.PictureBox Picture3 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   255
            Left            =   0
            ScaleHeight     =   255
            ScaleWidth      =   1455
            TabIndex        =   44
            Top             =   600
            Width           =   1455
            Begin VB.CheckBox chkFields 
               Alignment       =   1  'Right Justify
               Caption         =   "Posted:"
               DataField       =   "BANK TRANS Posted YN"
               Height          =   255
               Index           =   9
               Left            =   120
               TabIndex        =   45
               Top             =   0
               Width           =   1335
            End
         End
         Begin VB.CheckBox chkFields 
            Alignment       =   1  'Right Justify
            Caption         =   "This is a beginning balance:"
            DataField       =   "BANK TRANS Beg Balance"
            Height          =   255
            Index           =   0
            Left            =   2520
            TabIndex        =   43
            Top             =   600
            Width           =   2415
         End
         Begin VB.Label lblLabels 
            Caption         =   "Notes:"
            Height          =   255
            Index           =   14
            Left            =   120
            TabIndex        =   50
            Top             =   0
            Width           =   1455
         End
      End
      Begin VB.PictureBox pcMajor 
         BorderStyle     =   0  'None
         Height          =   5175
         Left            =   360
         ScaleHeight     =   5175
         ScaleWidth      =   7575
         TabIndex        =   14
         Top             =   480
         Width           =   7575
         Begin VB.CommandButton cmdlookupReceipt 
            Height          =   285
            Left            =   3000
            Picture         =   "frm_Bank_Transaction.frx":134D
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   2280
            Width           =   375
         End
         Begin VB.CommandButton cmdLookupDepo 
            Height          =   285
            Left            =   3000
            Picture         =   "frm_Bank_Transaction.frx":1657
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   1560
            Width           =   375
         End
         Begin VB.CommandButton cmdDate 
            Height          =   285
            Index           =   0
            Left            =   3000
            Picture         =   "frm_Bank_Transaction.frx":1961
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   840
            Width           =   375
         End
         Begin VB.CommandButton cmdLookupBT 
            Height          =   285
            Left            =   3000
            Picture         =   "frm_Bank_Transaction.frx":1C6B
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox txtFields 
            DataField       =   "BANK TRANS ID"
            Height          =   285
            Index           =   0
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   120
            Width           =   1455
         End
         Begin VB.TextBox txtFields 
            DataField       =   "BANK TRANS Ext Document No"
            Height          =   285
            Index           =   1
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   21
            Top             =   480
            Width           =   1455
         End
         Begin VB.TextBox txtFields 
            DataField       =   "BANK TRANS Amount"
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
            Index           =   2
            Left            =   1560
            TabIndex        =   20
            Top             =   1920
            Width           =   1815
         End
         Begin VB.TextBox txtFields 
            DataField       =   "BANK TRANS Date"
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
            Index           =   3
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   19
            Top             =   840
            Width           =   1455
         End
         Begin VB.TextBox txtFields 
            DataField       =   "BANK TRANS Bank Acct 1"
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
            Index           =   4
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   18
            Top             =   1560
            Width           =   1455
         End
         Begin VB.TextBox txtFields 
            DataField       =   "BANK TRANS Reference"
            Height          =   285
            Index           =   5
            Left            =   1560
            TabIndex        =   17
            Top             =   1200
            Width           =   1815
         End
         Begin VB.TextBox txtFields 
            DataField       =   "BANK TRANS Bank Acct 2"
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
            Index           =   12
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   2280
            Width           =   1455
         End
         Begin VB.ComboBox cbfields 
            DataField       =   "BANK TRANS Type"
            Height          =   315
            ItemData        =   "frm_Bank_Transaction.frx":1F75
            Left            =   5160
            List            =   "frm_Bank_Transaction.frx":1F88
            TabIndex        =   15
            Text            =   "Combo1"
            Top             =   480
            Width           =   2415
         End
         Begin MSDataGridLib.DataGrid grdDatagrid 
            Bindings        =   "frm_Bank_Transaction.frx":1FCB
            Height          =   2415
            Left            =   0
            TabIndex        =   33
            Top             =   2760
            Width           =   7575
            _ExtentX        =   13361
            _ExtentY        =   4260
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
            ColumnCount     =   7
            BeginProperty Column00 
               DataField       =   "AR PAY Reconciled"
               Caption         =   "Post"
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
               DataField       =   "AR PAY Type"
               Caption         =   "Type"
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
               DataField       =   "AR PAY Customer No"
               Caption         =   "Customer No"
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
               DataField       =   "AR PAY Check No"
               Caption         =   "Check No"
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
               DataField       =   "AR PAY Transaction Date"
               Caption         =   "Trans. Date"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "M/d/yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column05 
               DataField       =   "AR PAY Amount"
               Caption         =   "Amount"
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
               DataField       =   "AR PAY NSF"
               Caption         =   "NSF"
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
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
                  ColumnWidth     =   510.236
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1140.095
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   1214.929
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   1244.976
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   1035.213
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   1244.976
               EndProperty
               BeginProperty Column06 
                  Alignment       =   2
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   615.118
               EndProperty
            EndProperty
         End
         Begin VB.Frame frtype 
            Height          =   1815
            Left            =   5160
            TabIndex        =   27
            Top             =   720
            Width           =   2415
            Begin VB.PictureBox Picture1 
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   1335
               Left            =   360
               ScaleHeight     =   1335
               ScaleWidth      =   1575
               TabIndex        =   28
               Top             =   360
               Width           =   1575
               Begin VB.OptionButton optSelect 
                  Caption         =   "Withdrawal"
                  Height          =   255
                  Index           =   0
                  Left            =   0
                  TabIndex        =   32
                  Top             =   0
                  Width           =   1455
               End
               Begin VB.OptionButton optSelect 
                  Caption         =   "Deposit"
                  Height          =   255
                  Index           =   1
                  Left            =   0
                  TabIndex        =   31
                  Top             =   360
                  Width           =   1455
               End
               Begin VB.OptionButton optSelect 
                  Caption         =   "Transfer"
                  Height          =   255
                  Index           =   2
                  Left            =   0
                  TabIndex        =   30
                  Top             =   720
                  Width           =   1455
               End
               Begin VB.OptionButton optSelect 
                  Caption         =   "Deposit Slip"
                  Height          =   255
                  Index           =   3
                  Left            =   0
                  TabIndex        =   29
                  Top             =   1080
                  Width           =   1455
               End
            End
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Transaction ID:  "
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   41
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Document No:  "
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   40
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Amount:  "
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   39
            Top             =   1920
            Width           =   1335
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Transaction Date:  "
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   38
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Deposit To:  "
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   37
            Top             =   1560
            Width           =   1335
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Reference:  "
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   36
            Top             =   1200
            Width           =   1335
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Transaction Type:"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   6
            Left            =   5160
            TabIndex        =   35
            Top             =   240
            Width           =   2415
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Receipt Account:  "
            Height          =   255
            Index           =   12
            Left            =   240
            TabIndex        =   34
            Top             =   2280
            Width           =   1335
         End
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
      ScaleWidth      =   8205
      TabIndex        =   6
      Top             =   7695
      Width           =   8205
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Picture         =   "frm_Bank_Transaction.frx":1FE0
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Picture         =   "frm_Bank_Transaction.frx":2322
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Picture         =   "frm_Bank_Transaction.frx":2664
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Picture         =   "frm_Bank_Transaction.frx":29A6
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   11
         Top             =   0
         Width           =   3360
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
      ScaleWidth      =   8205
      TabIndex        =   0
      Top             =   7395
      Width           =   8205
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Ca&ncel"
         Height          =   300
         Left            =   5400
         TabIndex        =   12
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   300
         Left            =   4320
         TabIndex        =   2
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   300
         Left            =   3240
         TabIndex        =   3
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   300
         Left            =   2160
         TabIndex        =   4
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   300
         Left            =   1080
         TabIndex        =   1
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   300
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bank Transaction"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   51
      Top             =   120
      Width           =   7305
   End
End
Attribute VB_Name = "frm_Bank_Transaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ADOprimaryrs As ADODB.Recordset
Attribute ADOprimaryrs.VB_VarHelpID = -1
Dim SQLExecute As ADODB.Recordset

Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean

Dim db As ADODB.Connection

Private Sub CbSelection(cbString As String, Optional ReceiptType As String)
        lblLabels(5).Caption = "Reference"
        lblLabels(4).Caption = "Bank Account"
        lblLabels(2).Caption = "Amount"
        lblLabels(12).Caption = "GL Account"

Select Case cbString
    Case "Withdrawal"
        optSelect(0).Value = True
        chkFields(0).Visible = False
        grddatagrid.Visible = False
        cmdPost.Caption = "&Post"
        cmdPost.ToolTipText = "Post Bank Withdrawal"
        pcMajor.Height = 2655
    Case "Transfer To", "Transfer From"
        optSelect(2).Value = True
        chkFields(0).Visible = False
        grddatagrid.Visible = False
        cmdPost.Caption = "&Post"
        cmdPost.ToolTipText = "Post Bank Transfer"
        pcMajor.Height = 2655
    Case "Deposit"
        optSelect(1).Value = True
        chkFields(0).Visible = True
        grddatagrid.Visible = False
        cmdPost.Caption = "&Post"
        cmdPost.ToolTipText = "Post Bank Deposit"
        pcMajor.Height = 2655
    Case "Deposit Slip"
        optSelect(3).Value = True
        lblLabels(4).Caption = "Deposit To"
        lblLabels(12).Caption = "Receipt Acct"
        chkFields(0).Visible = True
        'If ReceiptType = "" Then
        '    grddatagrid.Visible = False
        '    pcMajor.Height = 2655
        'Else
            grddatagrid.Visible = True
            grdDataSource ReceiptType
            pcMajor.Height = 5295
        'End If
        cmdPost.Caption = "&Save"
        cmdPost.ToolTipText = "Save Deposit Slip"
End Select
        
        Form_Resize
End Sub

Private Sub cbfields_Change()
    CbSelection cbfields.Text, txtfields(12)
End Sub

Private Sub cbFields_Click()
    CbSelection cbfields.Text, txtfields(12)
End Sub

Private Sub cbfields_LostFocus()
    If CheckCombo(cbfields) Then
        MsgBox "Please make a new selection", vbCritical, "Error"
        cbfields.SetFocus
    End If
End Sub

Private Sub cmdDate_Click(Index As Integer)
    Menu_Calendar.WhoCallMe True, 1322
    'Menu_Calendar.Show vbModal
    txtfields(3).SetFocus
End Sub

Private Sub cmdLookupBT_Click()
If mbAddNewFlag = True Then Exit Sub
   AllLookup.ToWhichRecord ADOprimaryrs, "GL Entry", "ID//Type//Date//Reference"
   'AllLookup.Show vbModal
   txtfields(1).SetFocus
End Sub

Private Sub cmdLookupDepo_Click()
   AllLookup.GetWhichTable 1460, "Select [GL COA Account No],[GL COA Account Name]," & _
   "[GL COA Asset Type] From [GL Chart Of Accounts] ", "GL Accounts", _
   "Account No//Account Type//Account Type", db
   'AllLookup.Show vbModal
   txtfields(4).SetFocus
End Sub

Private Sub cmdlookupReceipt_Click()
   AllLookup.GetWhichTable 1465, "Select [GL COA Account No],[GL COA Account Name]," & _
   "[GL COA Asset Type] From [GL Chart Of Accounts] ", "GL Accounts", _
   "Account No//Account Type//Account Type", db
   'AllLookup.Show vbModal
   txtfields(12).SetFocus
   CbSelection cbfields.Text, txtfields(12)
End Sub

Private Sub cmdPost_Click()

'  On Error GoTo cmdPost_Click_Error
'    Me.Refresh
  If txtfields(0) = "" Then Exit Sub
  If txtfields(2) = "$0.00" Then
    MsgBox "Can't post a transaction with zero amount"
    Exit Sub
  End If
  'Post this transaction to the general ledger
   Dim Success%

  'Force record save
  'Cmd.RunMacro "Save Record"
  ADOprimaryrs.Update
  If CheckEmpty() = False Then Exit Sub

  ShowStatus False

  db.BeginTrans
    Select Case optType
    Case 1
      Success% = PostWithdrawal(CLng(txtfields(0)))
    Case 2
      Success% = PostDeposit(CLng(txtfields(0)))
    Case 3
      Success% = PostTransfer(CLng(txtfields(0)))
    Case 4
      'Success% = PostDepositSlip(CLng(txtFields(0)))
      Success% = True
    End Select
    If Success% = False Then
      db.RollbackTrans
      MsgBox "Transaction NOT Posted."
    Else
      db.CommitTrans
      MsgBox "Transaction Posted."
      'Dim rs As ADODB.Recordset
      'Dim TempID&
      'TempID& = adoPrimaryRS![BANK TRANS ID]
      'Me.Requery
      'Set rs = New ADODB.Recordset
      'Set rs = adoPrimaryRS.Clone
      'rs.FindFirst "[BANK TRANS ID] = " & TempID&
        'rs.Find "[BANK TRANS ID] = " & TempID&
        'chkFields(9).Value = 1
        ADOprimaryrs("BANK TRANS Posted YN") = True
      ADOprimaryrs.Update
      'DoCmd.GoToRecord A_FORM, "Bank Transactions", A_NEWREC
      'DoCmd.GoToControl "BANK TRANS Ext Document No"
    End If
    cmdPost.Enabled = False
    cmdClose.Enabled = True
    cmdFirst.Enabled = True
    cmdPrevious.Enabled = True
    cmdLast.Enabled = True
    cmdNext.Enabled = True

  ShowStatus False
  
  Exit Sub
  
RecordLocked:
  db.RollbackTrans
  Exit Sub

UnableToPost:
  db.RollbackTrans
  Exit Sub

cmdPost_Click_Error:
  Call ErrorLog("Bank Transactions", "cmdPost_Click", Now, Err.Number, Err.Description, True, db)
  Resume Next

End Sub


Private Sub SelectAll_Click()
    PostStatus False
End Sub

Private Sub UnSelectAll_Click()
    PostStatus True
End Sub

Private Sub Form_Load()
ShowStatus True
On Error GoTo FormErr
  Set db = New ADODB.Connection
  db.CursorLocation = adUseClient
  db.Open gblADOProvider

  Set ADOprimaryrs = New ADODB.Recordset
  Set adoPrimaryRS2 = New ADODB.Recordset
  
  OpenDB
  
ShowStatus False
mbDataChanged = False
Exit Sub
FormErr:
  MsgBox Err.Description
  ShowStatus False
End Sub

Private Sub OpenDB()
  ADOprimaryrs.Open "SELECT [BANK TRANS Ext Document No],[BANK TRANS Date]," & _
  "[BANK TRANS Reference],[BANK TRANS Amount],[BANK TRANS ID],[BANK TRANS Bank Acct 1]," & _
  "[BANK TRANS Bank Acct 2],[BANK TRANS Notes],[BANK TRANS Type],[BANK TRANS Beg Balance]," & _
  "[BANK TRANS Posted YN] FROM [BANK Transaction]", db, adOpenKeyset, adLockOptimistic, adCmdText
      
  Dim Ctrl As Control
  For Each Ctrl In Me.Controls
    If TypeOf Ctrl Is TextBox Or TypeOf Ctrl Is CheckBox Or TypeOf Ctrl Is ComboBox Then
        Set Ctrl.DataSource = ADOprimaryrs
        If TypeOf Ctrl Is TextBox And Ctrl.DataField <> "" Then
           If ADOprimaryrs("" & Ctrl.DataField & "").Type = 202 Then Ctrl.MaxLength = ADOprimaryrs("" & Ctrl.DataField & "").DefinedSize
        End If
    End If
  Next
  
  If CheckNewDB(ADOprimaryrs, "Bank Transaction") = True Then
    cmdAdd_Click
  End If
  'Set grdDataGrid.DataSource = adoPrimaryRS2
  grddatagrid.Columns(0).Button = True
  
GetTextColor Me
End Sub

Private Sub Form_Resize()
  If fMainForm.WindowState = 1 Then Exit Sub
  If Me.WindowState = 0 Then
  ElseIf Me.WindowState = 2 Then
    GoTo SkipResize
  Else
    Exit Sub
  End If
  
  Me.Width = 8325
  
SkipResize:
  If pcMajor.Height = 2655 Then
    picMinor.Top = 3360
    frPrimary.Height = 4575
    UnselectAll.Visible = False
    SelectAll.Visible = False
    If Me.WindowState = 0 Then Me.Height = 6120
  Else
    picMinor.Top = 5760
    frPrimary.Height = 6855
    UnselectAll.Visible = True
    SelectAll.Visible = True
    If Me.WindowState = 0 Then Me.Height = 8370
  End If
  
  frPrimary.Left = (Me.ScaleWidth - frPrimary.Width) / 2
  Label1.Left = frPrimary.Left
  Label1.Width = frPrimary.Width
  frPrimary.Top = (Me.ScaleHeight - frPrimary.Height - picStatBox.Height - picButtons.Height) / 2 + 230
  
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
    ShowStatus False
    If UnloadForm(ADOprimaryrs) = 0 Then
        db.Close
        Set db = Nothing
    Else
        Cancel = 1
    End If
    'SQLExecute.Close
    Set SQLExecute = Nothing
Exit Sub
FormErr:
  MsgBox Err.Description
  ShowStatus False
End Sub

Private Sub adoPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
Dim FlagBoolean As Boolean
If ADOprimaryrs.BOF Or ADOprimaryrs.EOF Then Exit Sub
If mbAddNewFlag = True Then
    picPosted.Visible = False
    pcMajor.Enabled = True
    Exit Sub
End If
If ADOprimaryrs![BANK TRANS Posted YN] = True Then
   picPosted.Visible = True
   FlagBoolean = False
Else
   picPosted.Visible = False
   FlagBoolean = True
End If
   pcMajor.Enabled = FlagBoolean
   cmdPost.Enabled = FlagBoolean
   cmdDelete.Enabled = FlagBoolean
   
  
  'This will display the current record position for this recordset
  lblStatus.Caption = "Record: " & CStr(ADOprimaryrs.AbsolutePosition) & " of " & CStr(ADOprimaryrs.RecordCount)
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
  'On Error GoTo AddErr
ShowStatus True
If cmdAdd.Caption = "&Save" Then
     If Not CheckEmpty Then
        ShowStatus False
        Exit Sub
     End If
     With ADOprimaryrs
         .UpdateBatch adAffectAll
         .MovePrevious
         'grdDatagrid.HoldFields
         'grdDatagrid.ReBind
         'grdDatagrid.Refresh
         'NewLoad = False
         ADOprimaryrs.Requery
         '.Find "[BANK TRANS Ext Document No]='" & "TRANS " & AppLoginName & "'"
         'txtFields(1).SetFocus
         'txtFields(1) = "TRANS " & ![BANK TRANS ID] + 1000
         '.UpdateBatch adAffectAll
     End With
     cmdAdd.Caption = "&Add"
     mbAddNewFlag = False
     txtfields(2).Locked = True
     SetButtons True
     cmdPost.Enabled = True
Else
  With ADOprimaryrs
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
     'NewLoad = True
     mbAddNewFlag = True
     .AddNew
     lblStatus.Caption = "Add record"
     cmdPost.Enabled = False
     txtfields(1) = "TRANS " & AppLoginName
     txtfields(14) = "-Created by " & AppLoginName
     txtfields(2) = "$0.00"
     chkFields(9).Value = 0
     txtfields(3).Text = FormatDate(Now)
     SetButtons False
     pcMajor.Height = 2655
     Form_Resize
  End With
  cmdAdd.Caption = "&Save"
  txtfields(2).Locked = False
End If
  GetTextColor Me
  ShowStatus False
Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
     If Not DataDelete(ADOprimaryrs, Me, True) Then
     End If
End Sub

Private Sub cmdRefresh_Click()
  'This is only needed for multi user apps
  'On Error GoTo RefreshErr
  'Set grdDataGrid.DataSource = Nothing
  ADOprimaryrs.Update
  ADOprimaryrs.Requery
  'Set grdDataGrid.DataSource = adoPrimaryRS("ChildCMD").UnderlyingValue
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub

Private Sub cmdEdit_Click()
  'On Error GoTo EditErr

  lblStatus.Caption = "Edit record"
  mbEditFlag = True
  SetButtons True
  Exit Sub

EditErr:
  MsgBox Err.Description
End Sub
Private Sub cmdCancel_Click()
On Error GoTo FormErr
  ShowStatus True
  SetButtons True
  cmdAdd.Caption = "&Add"
  cmdCancel.Visible = False
  mbEditFlag = False
  mbAddNewFlag = False
  ADOprimaryrs.CancelUpdate
  ADOprimaryrs.CancelBatch
  'NewLoad = False
  cmdPost.Enabled = True
  If ADOprimaryrs.RecordCount > 0 Then
    'MsgBox ADOprimaryrs.EditMode
    ADOprimaryrs.MoveLast
  Else
    MsgBox "No data to publish. Exiting " & Me.Caption
    Unload Me
    Exit Sub
  End If
  ADOprimaryrs.Resync adAffectCurrent
  If mvBookMark > 0 Then
    ADOprimaryrs.Bookmark = mvBookMark
  Else
    ADOprimaryrs.MoveFirst
  End If
  mbDataChanged = False
  txtfields(2).Locked = True
  GetTextColor Me
  ShowStatus False
Exit Sub
FormErr:
  MsgBox Err.Description
  Resume Next
End Sub

Private Sub cmdUpdate_Click()
'Dim FlagStatus As Boolean
    
  'FlagStatus = False

  Call UpdateButton(ADOprimaryrs, mbAddNewFlag)
  
  'mbEditFlag = Not FlagStatus
  
  'SetButtons FlagStatus
  txtfields(2).Locked = True
  GetTextColor Me
  'mbDataChanged = Not FlagStatus
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
    
    If mbAddNewFlag = True Then
        cmdAdd.Caption = "&Save"
        cmdCancel.Visible = True
        cmdCancel.Left = cmdUpdate.Left
        cmdCancel.Top = cmdUpdate.Top
    Else
        cmdAdd.Visible = bVal
        cmdCancel.Visible = False
    End If
        cmdUpdate.Visible = bVal
        cmdDelete.Visible = bVal
        cmdClose.Visible = bVal
        cmdRefresh.Visible = bVal
        cmdNext.Enabled = bVal
        cmdFirst.Enabled = bVal
        cmdLast.Enabled = bVal
        cmdPrevious.Enabled = bVal
End Sub

Private Function CheckEmpty() As Boolean
 For Each Ctrl In Me.txtfields
    If Ctrl.Text = "" Then
      Select Case Ctrl.Index
      Case 0
      Case Else
        MsgBox "There is an empty data in " & lblLabels(Ctrl.Index), vbInformation, "Empty Data"
        CheckEmpty = False
        Exit Function
      End Select
    End If
 Next
 If cbfields.Text = "" Or cbfields.Text = "Deposit Slip" Then
    cbfields.Text = cbfields.List(1)
    cbFields_Click
 End If
 CheckEmpty = True
End Function

Private Sub grdDataSource(Acct2 As String)
Dim SQLstatement As String

SQLstatement = "SELECT  [AR PAY Reconciled],[AR PAY Type],[AR PAY Customer No], " & _
"[AR PAY Check No],[AR PAY Transaction Date],[AR PAY Amount], [AR PAY NSF] " & _
"From [AR Payment Header] WHERE [AR PAY Bank Account]='" & Acct2 & "' " & _
"AND [AR PAY Type]='payment' AND [AR PAY Reconciled]=False AND [AR PAY Posted YN]=True " & _
"AND [AR PAY Deposited YN]=False ORDER BY [AR PAY Transaction Date]"

Set SQLExecute = New ADODB.Recordset
SQLExecute.Open SQLstatement, db, adOpenStatic, adLockOptimistic, adCmdText
Set grddatagrid.DataSource = SQLExecute
End Sub

Private Sub PostStatus(PstStatus As Boolean)
Dim i As Integer
    For i = 0 To SQLExecute.RecordCount - 1
        grddatagrid.Row = i
      If PstStatus = True Then
        If grddatagrid.Columns(0) = "Yes" Then grddatagrid.Columns(0) = "No"
      Else
        If grddatagrid.Columns(0) = "No" Then grddatagrid.Columns(0) = "Yes"
      End If
            SQLExecute.Update
            CalcTotals
    Next
End Sub

Private Sub grdDataGrid_ButtonClick(ByVal ColIndex As Integer)
    If grddatagrid.Row = -1 Or grddatagrid.Columns(0) = "" Then Exit Sub
         SendKeys ("{ENTER}")
   If grddatagrid.Columns(0).Text = "No" Then
      grddatagrid.Columns(0).Text = "Yes"
   Else
      grddatagrid.Columns(0).Text = "No"
   End If
         SendKeys ("{ENTER}")
         SendKeys ("{down}")
         SendKeys ("{up}")
         CalcTotals
End Sub

Private Sub CalcTotals()
Dim TempCount As Currency
  With SQLExecute
    If .RecordCount = 0 Then
      txtfields(2) = "$0.00"
      Exit Sub
    End If
    .MoveFirst
    Do While Not .EOF
      If ![AR PAY Reconciled] = True Then
        TempCount = TempCount + ![AR PAY Amount]
      End If
        .MoveNext
    Loop
  End With
  txtfields(2).SetFocus
  txtfields(2) = FormatCurr(TempCount)
  'Dim AllUpdate As Boolean
  'If TempCount = 0 Then
  '  AllUpdate = False
  'Else
  '  AllUpdate = True
  'End If
  '  cmdClose.Enabled = AllUpdate
  '  cmdFirst.Enabled = AllUpdate
  '  cmdPrevious.Enabled = AllUpdate
  '  cmdLast.Enabled = AllUpdate
  '  cmdNext.Enabled = AllUpdate
End Sub

Private Sub txtFields_Change(Index As Integer)
If mbAddNewFlag = True Then Exit Sub
Select Case Index
Case 1
    If txtfields(1) = "TRANS " & AppLoginName Then
        ADOprimaryrs![BANK TRANS Ext Document No] = AppLoginName & Format(Now, "MMdd") & Format(ADOprimaryrs![BANK TRANS ID], "000")
        txtfields(1) = AppLoginName & Format(Now, "MMdd") & Format(ADOprimaryrs![BANK TRANS ID], "000")
        ADOprimaryrs.Update
    End If
End Select

End Sub

Private Sub txtFields_LostFocus(Index As Integer)
Select Case Index
Case 5
    If txtfields(Index).Text = "" Then
        MsgBox "You can leave this reference empty"
        txtfields(Index).Text = "Need a reference No"
    End If
Case 14
    If InStr(1, txtfields(Index), "-Created by " & AppLoginName, vbTextCompare) Then Exit Sub
    txtfields(Index) = txtfields(Index) & "-Created by " & AppLoginName
End Select


End Sub
