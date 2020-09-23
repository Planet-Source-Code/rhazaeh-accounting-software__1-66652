VERSION 5.00
Begin VB.Form frm_SYS_Setup_Accounting_Preferences 
   Caption         =   "Accounting Preferences"
   ClientHeight    =   6165
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   10245
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6165
   ScaleWidth      =   10245
   Begin VB.Frame frPrimary 
      Height          =   5295
      Left            =   0
      TabIndex        =   19
      Top             =   480
      Width           =   10215
      Begin VB.Frame Frame1 
         Height          =   2655
         Left            =   5280
         TabIndex        =   27
         Top             =   2520
         Width           =   4815
         Begin VB.CheckBox chkFields 
            Alignment       =   1  'Right Justify
            Caption         =   "Post to GL:"
            DataField       =   "SYS COM Post to GL"
            DataSource      =   "adoPrimaryRS"
            Height          =   255
            Left            =   3480
            TabIndex        =   9
            Top             =   600
            Width           =   1095
         End
         Begin VB.TextBox txtfieldsfr 
            DataField       =   "SYS COM GL Number"
            DataSource      =   "adoPrimaryRS"
            Height          =   285
            Index           =   4
            Left            =   1200
            TabIndex        =   8
            Top             =   600
            Width           =   735
         End
         Begin VB.Frame frPO 
            Caption         =   "Purchase Order Form"
            Height          =   1455
            Left            =   120
            TabIndex        =   30
            Top             =   1080
            Width           =   2295
            Begin VB.OptionButton optPO 
               Caption         =   "Use Plain Paper Form"
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   10
               Top             =   480
               Width           =   1935
            End
            Begin VB.OptionButton optPO 
               Caption         =   "Use Preprinted Form"
               Height          =   255
               Index           =   1
               Left            =   240
               TabIndex        =   11
               Top             =   840
               Width           =   1815
            End
            Begin VB.TextBox txtfieldsfr 
               DataField       =   "SYS COM PO Form"
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
               Index           =   2
               Left            =   1680
               TabIndex        =   31
               Top             =   120
               Visible         =   0   'False
               Width           =   615
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Invoice Form"
            Height          =   1455
            Left            =   2400
            TabIndex        =   28
            Top             =   1080
            Width           =   2295
            Begin VB.OptionButton optInvForm 
               Caption         =   "Use Plain Paper Form"
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   12
               Top             =   480
               Width           =   1935
            End
            Begin VB.OptionButton optInvForm 
               Caption         =   "Use Preprinted Form"
               Height          =   255
               Index           =   1
               Left            =   240
               TabIndex        =   13
               Top             =   840
               Width           =   1935
            End
            Begin VB.TextBox txtfieldsfr 
               DataField       =   "SYS COM Invoice Form"
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
               Index           =   3
               Left            =   1560
               TabIndex        =   29
               Top             =   120
               Visible         =   0   'False
               Width           =   615
            End
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "GL Number:"
            Height          =   255
            Index           =   4
            Left            =   200
            TabIndex        =   32
            Top             =   600
            Width           =   975
         End
      End
      Begin VB.TextBox txtfields 
         DataField       =   "SYS COM Bank Service Charges Acct"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   1
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   720
         Width           =   3015
      End
      Begin VB.CommandButton cmdAcct 
         Height          =   255
         Index           =   1
         Left            =   4320
         Picture         =   "frm_SYS_Setup_Accounting_Preferences.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtfields 
         DataField       =   "SYS COM Bank Interest Earned Acct"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   2
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   1440
         Width           =   3015
      End
      Begin VB.CommandButton cmdAcct 
         Height          =   255
         Index           =   2
         Left            =   4320
         Picture         =   "frm_SYS_Setup_Accounting_Preferences.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox txtfields 
         DataField       =   "SYS COM Retained Earnings Acct"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   3
         Left            =   6720
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   720
         Width           =   3015
      End
      Begin VB.CommandButton cmdAcct 
         Height          =   255
         Index           =   3
         Left            =   9720
         Picture         =   "frm_SYS_Setup_Accounting_Preferences.frx":0294
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   720
         Width           =   375
      End
      Begin VB.Frame frGL 
         Height          =   1095
         Left            =   120
         TabIndex        =   22
         Top             =   2520
         Width           =   4575
         Begin VB.OptionButton optGL 
            Caption         =   "System Date"
            Height          =   255
            Index           =   0
            Left            =   1440
            TabIndex        =   4
            Top             =   360
            Width           =   2295
         End
         Begin VB.OptionButton optGL 
            Caption         =   "Transaction Date"
            Height          =   255
            Index           =   1
            Left            =   1440
            TabIndex        =   5
            Top             =   720
            Width           =   2295
         End
         Begin VB.TextBox txtfieldsfr 
            DataField       =   "SYS COM GL Post By Date"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   3840
            TabIndex        =   23
            Top             =   120
            Visible         =   0   'False
            Width           =   615
         End
      End
      Begin VB.Frame frInv 
         Height          =   1095
         Left            =   120
         TabIndex        =   20
         Top             =   4080
         Width           =   4575
         Begin VB.OptionButton optInv 
            Caption         =   "Standard Cost"
            Height          =   255
            Index           =   0
            Left            =   1440
            TabIndex        =   6
            Top             =   360
            Width           =   2055
         End
         Begin VB.OptionButton optInv 
            Caption         =   "Average Cost"
            Height          =   255
            Index           =   1
            Left            =   1440
            TabIndex        =   7
            Top             =   720
            Width           =   2055
         End
         Begin VB.TextBox txtfieldsfr 
            DataField       =   "SYS COM Inventory Cost Method Last YN"
            BeginProperty DataFormat 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "True"
               FalseValue      =   "False"
               NullValue       =   "N/A"
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   7
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   3840
            TabIndex        =   21
            Top             =   120
            Visible         =   0   'False
            Width           =   615
         End
      End
      Begin VB.Label lblLabels 
         Caption         =   "Bank Charges"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   43
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblAcct 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   42
         Top             =   1020
         Width           =   4575
      End
      Begin VB.Label lblLabels 
         Caption         =   "Interest Earned"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   41
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label lblAcct 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   40
         Top             =   1740
         Width           =   4575
      End
      Begin VB.Label lblLabels 
         Caption         =   "Retained Earnings"
         Height          =   255
         Index           =   3
         Left            =   5280
         TabIndex        =   39
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblAcct 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   5280
         TabIndex        =   38
         Top             =   1020
         Width           =   4815
      End
      Begin VB.Label lblAcct 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bank GL Accounts"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   37
         Top             =   240
         Width           =   4575
      End
      Begin VB.Label lblAcct 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Year End Processing Accounts"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   5
         Left            =   5280
         TabIndex        =   36
         Top             =   240
         Width           =   4815
      End
      Begin VB.Label lblAcct 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "GL Post Date"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   35
         Top             =   2160
         Width           =   4575
      End
      Begin VB.Label lblAcct 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Inventory Costing Method"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   34
         Top             =   3720
         Width           =   4575
      End
      Begin VB.Label lblAcct 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "General Ledger Properties"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   8
         Left            =   5280
         TabIndex        =   33
         Top             =   2160
         Width           =   4815
      End
   End
   Begin VB.TextBox txtfieldsTemp 
      DataField       =   " "
      DataSource      =   "adoPrimaryRS"
      Height          =   285
      Left            =   10320
      TabIndex        =   17
      Top             =   1560
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   10245
      TabIndex        =   0
      Top             =   5865
      Width           =   10245
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   300
         Left            =   2160
         TabIndex        =   16
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   300
         Left            =   1080
         TabIndex        =   15
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   300
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.Label lblTop 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Accounting Preferences"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   44
      Top             =   120
      Width           =   4815
   End
   Begin VB.Label lblAcct 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   10320
      TabIndex        =   18
      Top             =   1860
      Visible         =   0   'False
      Width           =   3015
   End
End
Attribute VB_Name = "frm_SYS_Setup_Accounting_Preferences"
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

Private Sub cmdAcct_Click(Index As Integer)
    Dim SQLstatement As String
    Dim No As Integer
    Dim ghead As String
    Dim fhead As String
    
    txtFieldsTemp = ""
    
    No = 1650
    SQLstatement = "select [GL COA Account No], [GL COA Account Name]" & _
                    "from [GL Chart Of Accounts]"
    ghead = "Account Description"
    fhead = "Account No//Description"
    
    AllLookup.GetWhichTable No, SQLstatement, ghead, fhead, db
    'AllLookup.Show vbModal
    If txtFieldsTemp <> "" Then
        txtfields(Index) = txtFieldsTemp
        lblAcct(Index) = lblAcct(0)
        txtfields(Index).SetFocus
    End If
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

Private Sub Form_Load()
ShowStatus True
On Error GoTo FormErr
  Set db = New ADODB.Connection
  db.CursorLocation = adUseClient
  db.Open gblADOProvider
  
  OpenDB
  
  If CheckNewDB(ADOprimaryrs, "Accounting Preferences") = True Then
    ADOprimaryrs.AddNew
  'Else
  '  picLedger.BackColor = txtfieldsfr(5)
  End If
  
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
  ADOprimaryrs.Open "select [SYS COM Bank Service Charges Acct],[SYS COM Bank Interest Earned Acct]," & _
    "[SYS COM Retained Earnings Acct],[SYS COM GL Post By Date],[SYS COM Inventory Cost Method Last YN]," & _
    "[SYS COM Post to GL],[SYS COM GL Number],[SYS COM PO Form],[SYS COM Invoice Form]" & _
    "from [SYS Company]", db, adOpenStatic, adLockOptimistic, adCmdText

  Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtfields
    Set oText.DataSource = ADOprimaryrs
    If oText.Text <> "" Then lblAcct(oText.Index) = LookRecord("[GL COA Account Name]", "[GL Chart Of Accounts]", db, "[GL COA Account No] = '" & txtfields(oText.Index) & "'")
  Next
  
  Set chkFields.DataSource = ADOprimaryrs
  
  For Each oText In Me.txtfieldsfr
    Set oText.DataSource = ADOprimaryrs
  Next
  If ADOprimaryrs.RecordCount > 0 Then
      If txtfieldsfr(0).Text = 1 Then
        optGL(0).Value = True
      Else
        optGL(1).Value = True
      End If
    
      If txtfieldsfr(1).Text = True Then
        optInv(0).Value = True
      Else
        optInv(1).Value = True
      End If
    
      If txtfieldsfr(2).Text = 1 Then
        optPO(0).Value = True
      Else
        optPO(1).Value = True
      End If
    
      If txtfieldsfr(3).Text = 1 Then
        optInvForm(0).Value = True
      Else
        optInvForm(1).Value = True
      End If
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
  
  Me.Width = 10335
  Me.Height = 6540
  
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
Dim i As Integer
  For i = 1 To 3
    If txtfields(i) = "" Then
        MsgBox lblLabels(i) & " data must be filled", vbCritical, "Error"
        Cancel = 1
        Exit Sub
    End If
  Next
  'updates the checklist Accounting Preferences
  ShowStatus True
      EndLoad db, ADOprimaryrs, "Accounting Preferences"
      If ADOprimaryrs.RecordCount > 0 Then
        If ADOprimaryrs.EditMode <> 0 Then
          ADOprimaryrs.CancelUpdate
        End If
      End If
      ADOprimaryrs.Close
      Set ADOprimaryrs = Nothing
      db.Close
      Set db = Nothing
      Set frm_SYS_Setup_Accounting_Preferences = Nothing
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

Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr
  
  With ADOprimaryrs
    .UpdateBatch adAffectCurrent
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

Private Sub optGL_Click(Index As Integer)
Select Case Index
Case 0
    txtfieldsfr(0) = 1
Case 1
    txtfieldsfr(0) = 2
End Select
End Sub

Private Sub optInv_Click(Index As Integer)
Select Case Index
Case 0
    txtfieldsfr(1) = True
Case 1
    txtfieldsfr(1) = False
End Select
End Sub

Private Sub optInvForm_Click(Index As Integer)
Select Case Index
Case 0
    txtfieldsfr(3) = 1
Case 1
    txtfieldsfr(3) = 2
End Select
End Sub

Private Sub optPO_Click(Index As Integer)
Select Case Index
Case 0
    txtfieldsfr(2) = 1
Case 1
    txtfieldsfr(2) = 2
End Select
End Sub
