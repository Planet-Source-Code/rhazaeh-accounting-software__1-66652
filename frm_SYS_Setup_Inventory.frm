VERSION 5.00
Begin VB.Form frm_SYS_Setup_Inventory 
   Caption         =   "Inventory Setup"
   ClientHeight    =   3255
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   6525
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3255
   ScaleWidth      =   6525
   Begin VB.Frame frPrimary 
      Height          =   2415
      Left            =   0
      TabIndex        =   11
      Top             =   480
      Width           =   6495
      Begin VB.CheckBox chkFields 
         Alignment       =   1  'Right Justify
         Caption         =   "Last In First Out (Y/N)"
         DataField       =   "SYS COM Inventory Cost Method Last YN"
         DataSource      =   "adoPrimaryRS"
         Height          =   255
         Index           =   5
         Left            =   4320
         TabIndex        =   6
         Top             =   480
         Width           =   2055
      End
      Begin VB.TextBox txtFields 
         DataField       =   "SYS COM Items per Transaction"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   4
         Left            =   2520
         TabIndex        =   5
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox txtFields 
         DataField       =   "SYS COM Inventory Qty Digits"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   3
         Left            =   2520
         TabIndex        =   4
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox txtFields 
         DataField       =   "SYS COM Inventory Production Number"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   2
         Left            =   2520
         TabIndex        =   3
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox txtFields 
         DataField       =   "SYS COM Inventory Cost Digits"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   1
         Left            =   2520
         TabIndex        =   2
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txtFields 
         DataField       =   "SYS COM Inventory Adjustment Number"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   0
         Left            =   2520
         TabIndex        =   1
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Items Per Transaction"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   16
         Top             =   1920
         Width           =   2295
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Inventory Quantity Digits"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   15
         Top             =   1560
         Width           =   2295
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Inventory Production Number"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   14
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Inventory Cost Digits"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Inventory Adjustment Number"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   2295
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
      ScaleWidth      =   6525
      TabIndex        =   0
      Top             =   2955
      Width           =   6525
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   300
         Left            =   2160
         TabIndex        =   9
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   300
         Left            =   1080
         TabIndex        =   8
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   300
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.Label lblTop 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Inventory Setup"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "frm_SYS_Setup_Inventory"
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
'The recordset should only contain one record holding information pertaining to  a specific company
' inventory setup.

'Private Sub btInvAdjAcc_Click()
'    Dim No As Integer
'    Dim sql As String
'    Dim ghead As String
'    Dim fhead As String
    
'    No = 3
'    sql = "select [GL COA Account No], [GL COA Account Name]" & _
'        "from [GL Chart of Accounts]"
'    ghead = "Account Description"
'    fhead = "Account No//Description"
    
'    AllLookup.GetWhichTable No, sql, ghead, fhead,db
'    'AllLookup.Show vbModal
'    txtfields(0).SetFocus
'End Sub

'Private Sub btInvProdAcc_Click()
'    Dim No As Integer
'    Dim sql As String
'    Dim ghead As String
'    Dim fhead As String
    
'    No = 4
'    sql = "select [GL COA Account No], [GL COA Account Name]" & _
'        "from [GL Chart of Accounts]"
'    ghead = "Account Description"
'    fhead = "Account No//Description"
    
'    AllLookup.GetWhichTable No, sql, ghead, fhead,db
'    'AllLookup.Show vbModal
'    txtfields(2).SetFocus
'End Sub

Private Sub Form_Load()
On Error GoTo FormErr
  Set db = New ADODB.Connection
  db.CursorLocation = adUseClient
  db.Open gblADOProvider

  Set ADOprimaryrs = New ADODB.Recordset
  ADOprimaryrs.Open "select [SYS COM Inventory Adjustment Number],[SYS COM Inventory Cost Digits]," & _
    "[SYS COM Inventory Cost Method Last YN],[SYS COM Inventory Production Number]," & _
    "[SYS COM Inventory Qty Digits],[SYS COM Items per Transaction] from [SYS Company]", db, adOpenStatic, _
    adLockOptimistic

  Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtfields
    Set oText.DataSource = ADOprimaryrs
    If ADOprimaryrs("" & oText.DataField & "").Type = 202 Then oText.MaxLength = ADOprimaryrs("" & oText.DataField & "").DefinedSize
  Next

  Dim oCheck As CheckBox
    'Bind the Check boxes to the data provider
  For Each oCheck In Me.chkFields
    Set oCheck.DataSource = ADOprimaryrs
  Next
  
  If CheckNewDB(ADOprimaryrs, "Inventory") = True Then
    ADOprimaryrs.AddNew
  End If
  
  GetTextColor Me
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
  
  Me.Width = 6615
  Me.Height = 3630
  
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
      If ADOprimaryrs.RecordCount > 0 Then
        If ADOprimaryrs.EditMode <> 0 Then
          ADOprimaryrs.CancelUpdate
        End If
      End If
      ADOprimaryrs.Close
      Set ADOprimaryrs = Nothing
      db.Close
      Set db = Nothing
  Set frm_SYS_Setup_Inventory = Nothing
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
  'mbEditFlag = False
  'mbAddNewFlag = False
  'mbDataChanged = False
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

Private Sub txtfields_KeyPress(Index As Integer, KeyAscii As Integer)
    keyResponse = CtrlValidate(KeyAscii, "0123456789.")
    If keyResponse = True Then
    Else
       KeyAscii = 0
    End If
End Sub

Private Sub txtFields_LostFocus(Index As Integer)
    txtfields(Index) = Val(txtfields(Index))
End Sub
