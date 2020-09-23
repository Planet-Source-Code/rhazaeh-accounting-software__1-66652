VERSION 5.00
Begin VB.Form frm_SYS_Setup_Tax_Authorities 
   Caption         =   "Tax Authorities"
   ClientHeight    =   3180
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   5415
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3180
   ScaleWidth      =   5415
   Begin VB.Frame frprimary 
      Height          =   2055
      Left            =   0
      TabIndex        =   18
      Top             =   480
      Width           =   5415
      Begin VB.TextBox txtFields 
         DataField       =   "SYS TAX Account"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   10
         Left            =   2280
         TabIndex        =   5
         Top             =   1560
         Width           =   1695
      End
      Begin VB.TextBox txtFields 
         DataField       =   "SYS TAX Percent"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   7
         Left            =   2280
         TabIndex        =   4
         Top             =   1200
         Width           =   2055
      End
      Begin VB.TextBox txtFields 
         DataField       =   "SYS TAX Description"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   1
         Left            =   2280
         TabIndex        =   3
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox txtFields 
         DataField       =   "SYS TAX ID"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   0
         Left            =   2280
         TabIndex        =   1
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton btItemID 
         Height          =   285
         Left            =   3960
         Picture         =   "frm_SYS_Setup_Tax_Authorities.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton btItemCostOfSales 
         Height          =   285
         Left            =   3960
         Picture         =   "frm_SYS_Setup_Tax_Authorities.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Account"
         Height          =   255
         Index           =   18
         Left            =   600
         TabIndex        =   22
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Percent"
         Height          =   255
         Index           =   2
         Left            =   600
         TabIndex        =   21
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Description"
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   20
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Tax ID"
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   19
         Top             =   480
         Width           =   1575
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
      ScaleWidth      =   5415
      TabIndex        =   17
      Top             =   2580
      Width           =   5415
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   300
         Left            =   4320
         TabIndex        =   11
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   300
         Left            =   3240
         TabIndex        =   10
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   300
         Left            =   2160
         TabIndex        =   9
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   300
         Left            =   1080
         TabIndex        =   8
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
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
      ScaleWidth      =   5415
      TabIndex        =   0
      Top             =   2880
      Width           =   5415
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Picture         =   "frm_SYS_Setup_Tax_Authorities.frx":0294
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Picture         =   "frm_SYS_Setup_Tax_Authorities.frx":05D6
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Picture         =   "frm_SYS_Setup_Tax_Authorities.frx":0918
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Picture         =   "frm_SYS_Setup_Tax_Authorities.frx":0C5A
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   16
         Top             =   0
         Width           =   3360
      End
   End
   Begin VB.Label lblTop 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tax Authorities"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   23
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "frm_SYS_Setup_Tax_Authorities"
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
    txtfields(10).SetFocus
End Sub

Private Sub btItemID_Click()
    Dim ghead As String
    Dim fhead As String

    ghead = "Tax"
    fhead = "Tax ID//Description"
    AllLookup.ToWhichRecord ADOprimaryrs, ghead, fhead
    'AllLookup.Show vbModal
End Sub


Private Sub Form_Load()
ShowStatus True
On Error GoTo FormErr
    Set db = New ADODB.Connection
    db.CursorLocation = adUseClient
    db.Open gblADOProvider

  Dim sql As String
  sql = "select [SYS TAX ID],[SYS TAX Description],[SYS TAX Percent]," & _
    "[SYS TAX Account] from [SYS Tax] "
  Set ADOprimaryrs = New ADODB.Recordset
  ADOprimaryrs.Open sql, db, adOpenStatic, adLockOptimistic
  If ADOprimaryrs.RecordCount = 0 Then
    cmdUpdate.Enabled = False
    cmdDelete.Enabled = False
    cmdRefresh.Enabled = False
  End If
  
  Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtfields
    Set oText.DataSource = ADOprimaryrs
    If ADOprimaryrs("" & oText.DataField & "").Type = 202 Then oText.MaxLength = ADOprimaryrs("" & oText.DataField & "").DefinedSize
  Next
  
  'Lock these fields to avoid invalid entries
  txtfields(10).Locked = True
   
  If CheckNewDB(ADOprimaryrs, "Tax Group") = True Then
    cmdAdd_Click
  End If
  
  Me.Width = 5550
  Me.Height = 3540

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
  
  Me.Width = 5550
  Me.Height = 3540
  
SkipResize:
  frPrimary.Left = (Me.ScaleWidth - frPrimary.Width) / 2
  lblTop.Left = frPrimary.Left
  lblTop.Width = frPrimary.Width
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
  ShowStatus True
      EndLoad db, ADOprimaryrs, "Tax Authorities"
      If ADOprimaryrs.RecordCount > 0 Then
        If ADOprimaryrs.EditMode <> 0 Then
          ADOprimaryrs.CancelUpdate
        End If
      End If
      ADOprimaryrs.Close
      Set ADOprimaryrs = Nothing
      db.Close
      Set db = Nothing
Set frm_SYS_Setup_Tax_Authorities = Nothing
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

Private Sub cmdAdd_Click()
  'On Error GoTo AddErr
  With ADOprimaryrs
    If cmdAdd.Caption = "&Add" Then
        If Not (.BOF And .EOF) Then
            mvBookMark = .Bookmark
        End If
        .AddNew
        lblStatus.Caption = "Add record"
        mbAddNewFlag = True
        txtfields(0).Enabled = True
        'txtFields(0).SetFocus
        btItemID.Enabled = False
        cmdAdd.Caption = "&Cancel"
        cmdUpdate.Enabled = True
    Else
        mbAddNewFlag = False
        .CancelUpdate
        btItemID.Enabled = True
        txtfields(0).Enabled = False
        cmdAdd.Caption = "&Add"
        If RecordCount > 0 Then
            If mvBookMark > 0 Then
                .Bookmark = mvBookMark
            Else
                .MoveLast
            End If
        End If
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
     If Not DataDelete(ADOprimaryrs, Me, True) Then
     End If
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
    If Trim(txtfields(0).Text) = "" Then
        MsgBox lblLabels(0).Caption & " must be filled", vbCritical, "Error"
        Exit Sub
    End If
    With ADOprimaryrs
        If .RecordCount = 0 Then Exit Sub 'no records to update
        .Update
        If mbAddNewFlag Then 'requery to get default value assigned by database
            .Requery
            .MoveLast
            mbAddNewFlag = False
        End If
        btItemID.Enabled = True
        'reenable the necessary buttons
        cmdAdd.Caption = "&Add"
        txtfields(0).Enabled = False
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
