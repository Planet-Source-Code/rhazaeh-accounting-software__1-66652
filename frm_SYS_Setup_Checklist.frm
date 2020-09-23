VERSION 5.00
Begin VB.Form frm_SYS_Setup_Checklist 
   Caption         =   "Setup Checklist"
   ClientHeight    =   4350
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   10140
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4350
   ScaleWidth      =   10140
   Begin VB.Frame frPrimary 
      Height          =   3495
      Left            =   0
      TabIndex        =   21
      Top             =   480
      Width           =   10095
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   4920
         Top             =   1440
      End
      Begin VB.CommandButton cmdCheckList 
         Caption         =   ">"
         Height          =   255
         Index           =   15
         Left            =   2880
         TabIndex        =   4
         Top             =   1800
         Width           =   255
      End
      Begin VB.CommandButton cmdCheckList 
         Caption         =   ">"
         Height          =   255
         Index           =   16
         Left            =   6000
         TabIndex        =   8
         Top             =   720
         Width           =   255
      End
      Begin VB.CommandButton cmdCheckList 
         Caption         =   ">"
         Height          =   255
         Index           =   13
         Left            =   2880
         TabIndex        =   5
         Top             =   2160
         Width           =   255
      End
      Begin VB.CommandButton cmdCheckList 
         Caption         =   ">"
         Height          =   255
         Index           =   12
         Left            =   6000
         TabIndex        =   13
         Top             =   2520
         Width           =   255
      End
      Begin VB.CommandButton cmdCheckList 
         Caption         =   ">"
         Height          =   255
         Index           =   11
         Left            =   6000
         TabIndex        =   14
         Top             =   2880
         Width           =   255
      End
      Begin VB.CommandButton cmdCheckList 
         Caption         =   ">"
         Height          =   255
         Index           =   8
         Left            =   2880
         TabIndex        =   6
         Top             =   2520
         Width           =   255
      End
      Begin VB.CommandButton cmdCheckList 
         Caption         =   ">"
         Height          =   255
         Index           =   7
         Left            =   6000
         TabIndex        =   11
         Top             =   1800
         Width           =   255
      End
      Begin VB.CommandButton cmdCheckList 
         Caption         =   ">"
         Height          =   255
         Index           =   6
         Left            =   2880
         TabIndex        =   7
         Top             =   2880
         Width           =   255
      End
      Begin VB.CommandButton cmdCheckList 
         Caption         =   ">"
         Height          =   255
         Index           =   5
         Left            =   6000
         TabIndex        =   9
         Top             =   1080
         Width           =   255
      End
      Begin VB.CommandButton cmdCheckList 
         Caption         =   ">"
         Height          =   255
         Index           =   4
         Left            =   6000
         TabIndex        =   10
         Top             =   1440
         Width           =   255
      End
      Begin VB.CommandButton cmdCheckList 
         Caption         =   ">"
         Height          =   255
         Index           =   2
         Left            =   2880
         TabIndex        =   2
         Top             =   1080
         Width           =   255
      End
      Begin VB.CommandButton cmdCheckList 
         Caption         =   ">"
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   1
         Top             =   720
         Width           =   255
      End
      Begin VB.CommandButton cmdCheckList 
         Caption         =   ">"
         Height          =   255
         Index           =   0
         Left            =   6000
         TabIndex        =   12
         Top             =   2160
         Width           =   255
      End
      Begin VB.CommandButton cmdCheckList 
         Caption         =   ">"
         Height          =   255
         Index           =   14
         Left            =   2880
         TabIndex        =   3
         Top             =   1440
         Width           =   255
      End
      Begin VB.CommandButton cmdCheckList 
         Caption         =   ">"
         Height          =   255
         Index           =   3
         Left            =   9240
         TabIndex        =   15
         Top             =   720
         Width           =   255
      End
      Begin VB.CommandButton cmdCheckList 
         Caption         =   ">"
         Height          =   255
         Index           =   10
         Left            =   9240
         TabIndex        =   16
         Top             =   1080
         Width           =   255
      End
      Begin VB.CommandButton cmdCheckList 
         Caption         =   ">"
         Height          =   255
         Index           =   9
         Left            =   9240
         TabIndex        =   17
         Top             =   1440
         Width           =   255
      End
      Begin VB.Frame Frame1 
         Enabled         =   0   'False
         Height          =   2055
         Left            =   6480
         TabIndex        =   22
         Top             =   120
         Width           =   3495
         Begin VB.CheckBox chkFields 
            Alignment       =   1  'Right Justify
            Caption         =   "17.  Payroll:"
            DataField       =   "Payroll"
            DataSource      =   "adoPrimaryRS"
            Height          =   255
            Index           =   9
            Left            =   480
            TabIndex        =   25
            Top             =   1320
            Width           =   2175
         End
         Begin VB.CheckBox chkFields 
            Alignment       =   1  'Right Justify
            Caption         =   "16.  Projects:"
            DataField       =   "Projects"
            DataSource      =   "adoPrimaryRS"
            Height          =   255
            Index           =   10
            Left            =   480
            TabIndex        =   24
            Top             =   960
            Width           =   2175
         End
         Begin VB.CheckBox chkFields 
            Alignment       =   1  'Right Justify
            Caption         =   "15.  Credit Cards Used:"
            DataField       =   "Credit Cards Used"
            DataSource      =   "adoPrimaryRS"
            Height          =   255
            Index           =   3
            Left            =   480
            TabIndex        =   23
            Top             =   600
            Width           =   2175
         End
      End
      Begin VB.Frame frWajib 
         Enabled         =   0   'False
         Height          =   3255
         Left            =   120
         TabIndex        =   26
         Top             =   120
         Width           =   6375
         Begin VB.CheckBox chkFields 
            Alignment       =   1  'Right Justify
            Caption         =   "9.   Employees:"
            DataField       =   "Employees"
            DataSource      =   "adoPrimaryRS"
            Height          =   255
            Index           =   5
            Left            =   3360
            TabIndex        =   40
            Top             =   960
            Width           =   2415
         End
         Begin VB.CheckBox chkFields 
            Alignment       =   1  'Right Justify
            Caption         =   "8.   Vendors:"
            DataField       =   "Vendors"
            DataSource      =   "adoPrimaryRS"
            Height          =   255
            Index           =   16
            Left            =   3360
            TabIndex        =   39
            Top             =   600
            Width           =   2415
         End
         Begin VB.CheckBox chkFields 
            Alignment       =   1  'Right Justify
            Caption         =   "10. Customers:"
            DataField       =   "Customers"
            DataSource      =   "adoPrimaryRS"
            Height          =   255
            Index           =   4
            Left            =   3360
            TabIndex        =   38
            Top             =   1320
            Width           =   2415
         End
         Begin VB.CheckBox chkFields 
            Alignment       =   1  'Right Justify
            Caption         =   "3.   Tax Authorities:"
            DataField       =   "Tax Authorities"
            DataSource      =   "adoPrimaryRS"
            Height          =   255
            Index           =   14
            Left            =   480
            TabIndex        =   37
            Top             =   1320
            Width           =   2175
         End
         Begin VB.CheckBox chkFields 
            Alignment       =   1  'Right Justify
            Caption         =   "7.   Item Categories:"
            DataField       =   "Item Categories"
            DataSource      =   "adoPrimaryRS"
            Height          =   255
            Index           =   6
            Left            =   480
            TabIndex        =   36
            Top             =   2760
            Width           =   2175
         End
         Begin VB.CheckBox chkFields 
            Alignment       =   1  'Right Justify
            Caption         =   "12. Accounting Preferences:"
            DataField       =   "Accounting Preferences"
            DataSource      =   "adoPrimaryRS"
            Height          =   255
            Index           =   0
            Left            =   3360
            TabIndex        =   35
            Top             =   2040
            Width           =   2415
         End
         Begin VB.CheckBox chkFields 
            Alignment       =   1  'Right Justify
            Caption         =   "4.   Tax Groups:"
            DataField       =   "Tax Groups"
            DataSource      =   "adoPrimaryRS"
            Height          =   255
            Index           =   15
            Left            =   480
            TabIndex        =   34
            Top             =   1680
            Width           =   2175
         End
         Begin VB.CheckBox chkFields 
            Alignment       =   1  'Right Justify
            Caption         =   "11. Items:"
            DataField       =   "Items"
            DataSource      =   "adoPrimaryRS"
            Height          =   255
            Index           =   7
            Left            =   3360
            TabIndex        =   33
            Top             =   1680
            Width           =   2415
         End
         Begin VB.CheckBox chkFields 
            Alignment       =   1  'Right Justify
            Caption         =   "13. Sales Preferences:"
            DataField       =   "Sales Preferences"
            DataSource      =   "adoPrimaryRS"
            Height          =   255
            Index           =   12
            Left            =   3360
            TabIndex        =   32
            Top             =   2400
            Width           =   2415
         End
         Begin VB.CheckBox chkFields 
            Alignment       =   1  'Right Justify
            Caption         =   "14. Purchasing Preferences:"
            DataField       =   "Purchasing Preferences"
            DataSource      =   "adoPrimaryRS"
            Height          =   255
            Index           =   11
            Left            =   3360
            TabIndex        =   31
            Top             =   2760
            Width           =   2415
         End
         Begin VB.CheckBox chkFields 
            Alignment       =   1  'Right Justify
            Caption         =   "5.   Shipping Methods:"
            DataField       =   "Shipping Methods"
            DataSource      =   "adoPrimaryRS"
            Height          =   255
            Index           =   13
            Left            =   480
            TabIndex        =   30
            Top             =   2040
            Width           =   2175
         End
         Begin VB.CheckBox chkFields 
            Alignment       =   1  'Right Justify
            Caption         =   "6.   Payment Terms:"
            DataField       =   "Payment Terms"
            DataSource      =   "adoPrimaryRS"
            Height          =   195
            Index           =   8
            Left            =   480
            TabIndex        =   29
            Top             =   2400
            Width           =   2175
         End
         Begin VB.CheckBox chkFields 
            Alignment       =   1  'Right Justify
            Caption         =   "2.   Company:"
            DataField       =   "Company"
            DataSource      =   "adoPrimaryRS"
            Height          =   255
            Index           =   2
            Left            =   480
            TabIndex        =   28
            Top             =   960
            Width           =   2175
         End
         Begin VB.CheckBox chkFields 
            Alignment       =   1  'Right Justify
            Caption         =   "1.   Chart Of Accounts:"
            DataField       =   "Chart Of Accounts"
            DataSource      =   "adoPrimaryRS"
            Height          =   255
            Index           =   1
            Left            =   480
            TabIndex        =   27
            Top             =   600
            Width           =   2175
         End
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"frm_SYS_Setup_Checklist.frx":0000
         Height          =   1215
         Left            =   6480
         TabIndex        =   41
         Top             =   2160
         Width           =   3495
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
      ScaleWidth      =   10140
      TabIndex        =   0
      Top             =   4050
      Width           =   10140
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   300
         Left            =   1080
         TabIndex        =   19
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   300
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.Label lblTop 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Setup Checklist"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   20
      Top             =   120
      Width           =   9840
   End
End
Attribute VB_Name = "frm_SYS_Setup_Checklist"
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

Private Sub cmdCheckList_Click(Index As Integer)
Select Case Index
Case 0
    frm_SYS_Setup_Accounting_Preferences.Show
Case 1
    frm_SYS_Setup_Chart_Of_Accounts.Show
Case 2
    frm_SYS_Setup_Company.Show
Case 3
    frm_LIST_Credit_Cards.Show
Case 4
    frm_AR_Customer.Show
Case 5
    frm_SYS_Setup_Employee.Show
Case 6
    frm_LIST_ALL_Types.ListType "frm_LIST_Item_Catagories"
Case 7
    frm_SYS_Setup_Items.Show
Case 8
    frm_LIST_Payment_Terms.Show
Case 9
    frm_SYS_Setup_Payroll.Show
Case 10
    frm_AR_Cust_Projects.Show
Case 11
    frm_SYS_Setup_Purchases.Show
Case 12
    frm_SYS_Setup_Sales.Show
Case 13
    frm_LIST_ALL_Types.ListType "frm_LIST_Shipping_Methods"
Case 14
    frm_SYS_Setup_Tax_Authorities.Show
Case 15
    frm_SYS_Setup_Tax_Group.Show
Case 16
    frm_AP_Vendor.Show
End Select
End Sub

' Note:
' This form only contains information pertaining to a specific company. The values should not be
' be editable. The fields should be check or unchecked depending on whether the related tables in
' the database is empty or not. The values of this table should be updated before unloading
' the related setup forms.
' yos

Private Sub Form_Load()
ShowStatus True
On Error GoTo FormErr
  Set db = New ADODB.Connection
  db.CursorLocation = adUseClient
  db.Open gblADOProvider

  OpenDB
  
  Me.Width = 10260
  Me.Height = 4755
  
  GetTextColor Me
  mbDataChanged = False
ShowStatus False
Exit Sub
FormErr:
  MsgBox Err.Description
  ShowStatus False
End Sub

Private Sub OpenDB()

  Set ADOprimaryrs = New ADODB.Recordset
  ADOprimaryrs.Open "select [Accounting Preferences],[Chart Of Accounts],Company," & _
    "[Credit Cards Used],Customers,Employees,[Item Categories],Items,[Payment Terms]," & _
    "Payroll,Projects,[Purchasing Preferences],[Sales Preferences],[Shipping Methods]," & _
    "[Tax Authorities],[Tax Groups],Vendors from [SYS Setup]", db, adOpenStatic, adLockReadOnly

    CallControlStatus
    Timer1.Enabled = True
End Sub

Private Sub CallControlStatus()
  Dim oCheck As CheckBox
    'Bind the Check boxes to the data provider
  For Each oCheck In Me.chkFields
    Set oCheck.DataSource = ADOprimaryrs
    If oCheck.Value = 1 Then
        cmdCheckList(oCheck.Index).Caption = "="
        'cmdCheckList(oCheck.Index).fo
    Else
        cmdCheckList(oCheck.Index).Caption = ">"
    End If
  Next
  IntroCheckList
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
  
  Me.Width = 10260
  Me.Height = 4755
  
SkipResize:
  frPrimary.Left = (Me.ScaleWidth - frPrimary.Width) / 2
  lblTop.Left = frPrimary.Left
  lblTop.Width = frPrimary.Width
  frPrimary.Top = (Me.ScaleHeight - frPrimary.Height - picButtons.Height) / 2 + 230
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If ADOprimaryrs.EditMode = True Then Exit Sub
  If KeyCode = vbKeyEscape Then cmdClose_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo FormErr
  ShowStatus True
      If ADOprimaryrs.RecordCount > 0 Then
        If ADOprimaryrs.EditMode <> 0 Then
          ADOprimaryrs.CancelUpdate
        End If
      End If
      Timer1.Enabled = False
      ADOprimaryrs.Close
      Set ADOprimaryrs = Nothing
      db.Close
      Set db = Nothing
      Set frm_SYS_Setup_Checklist = Nothing
  ShowStatus False
Exit Sub
FormErr:
  MsgBox Err.Description
  ShowStatus False
End Sub

Private Sub adoPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This will display the current record position for this recordset
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
  'On Error GoTo RefreshErr
  ADOprimaryrs.Requery
  CallControlStatus
  'OpenDB
  Exit Sub
RefreshErr:
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

Private Sub Timer1_Timer()
  ADOprimaryrs.Requery
  CallControlStatus
End Sub
