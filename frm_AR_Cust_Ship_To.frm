VERSION 5.00
Begin VB.Form frm_AR_Cust_Ship_To 
   Caption         =   "Maintain Ship To Data"
   ClientHeight    =   6555
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   6780
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6555
   ScaleWidth      =   6780
   Begin VB.Frame frPrimary 
      Height          =   5415
      Left            =   0
      TabIndex        =   28
      Top             =   480
      Width           =   6735
      Begin VB.TextBox txtFields 
         DataField       =   "AR SHIP State"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   7
         Left            =   3720
         TabIndex        =   9
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox txtFields 
         DataField       =   "AR SHIP Postal"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   8
         Left            =   4800
         TabIndex        =   10
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox txtFields 
         DataField       =   "AR SHIP Phone Ext"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   12
         Left            =   5040
         TabIndex        =   14
         Top             =   3240
         Width           =   615
      End
      Begin VB.TextBox txtFields 
         DataField       =   "AR SHIP Phone"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   11
         Left            =   2280
         TabIndex        =   13
         Top             =   3240
         Width           =   2175
      End
      Begin VB.TextBox txtFields 
         DataField       =   "AR SHIP Notes"
         DataSource      =   "adoPrimaryRS"
         Height          =   1095
         Index           =   14
         Left            =   2280
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Top             =   3960
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "AR SHIP Name"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   3
         Left            =   2280
         TabIndex        =   5
         Top             =   1080
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "AR SHIP ID"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   2
         Left            =   2280
         TabIndex        =   2
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtFields 
         DataField       =   "AR SHIP Fax"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   13
         Left            =   2280
         TabIndex        =   15
         Top             =   3600
         Width           =   3375
      End
      Begin VB.CheckBox chkFields 
         Alignment       =   1  'Right Justify
         Caption         =   "Default Ship To"
         DataField       =   "AR SHIP Default"
         DataSource      =   "adoPrimaryRS"
         Height          =   255
         Left            =   4200
         TabIndex        =   4
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtFields 
         DataField       =   "AR SHIP Customer ID"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   0
         Left            =   2280
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtFields 
         DataField       =   "AR SHIP Country"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   9
         Left            =   2280
         TabIndex        =   11
         Top             =   2520
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "AR SHIP Contact"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   10
         Left            =   2280
         TabIndex        =   12
         Top             =   2880
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "AR SHIP City"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   6
         Left            =   2280
         TabIndex        =   8
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox txtFields 
         DataField       =   "AR SHIP Address 2"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   5
         Left            =   2280
         TabIndex        =   7
         Top             =   1800
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "AR SHIP Address 1"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   4
         Left            =   2280
         TabIndex        =   6
         Top             =   1440
         Width           =   3375
      End
      Begin VB.CommandButton bt_Cust_ShipTo 
         Height          =   285
         Left            =   3600
         Picture         =   "frm_AR_Cust_Ship_To.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   720
         Width           =   375
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "State:"
         Height          =   255
         Index           =   14
         Left            =   3120
         TabIndex        =   41
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Zip:"
         Height          =   255
         Index           =   13
         Left            =   4320
         TabIndex        =   40
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Ext:"
         DataSource      =   "adoPrimaryRS"
         Height          =   255
         Index           =   12
         Left            =   4560
         TabIndex        =   39
         Top             =   3240
         Width           =   375
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Phone Number:"
         Height          =   255
         Index           =   11
         Left            =   840
         TabIndex        =   38
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Notes:"
         Height          =   255
         Index           =   10
         Left            =   840
         TabIndex        =   37
         Top             =   3960
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Ship To Name:"
         Height          =   255
         Index           =   9
         Left            =   840
         TabIndex        =   36
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Ship To ID:"
         Height          =   255
         Index           =   8
         Left            =   840
         TabIndex        =   35
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Fax Number:"
         Height          =   255
         Index           =   7
         Left            =   840
         TabIndex        =   34
         Top             =   3600
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Customer ID:"
         Height          =   255
         Index           =   5
         Left            =   840
         TabIndex        =   33
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Country:"
         Height          =   255
         Index           =   4
         Left            =   840
         TabIndex        =   32
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Contact Person:"
         Height          =   255
         Index           =   3
         Left            =   840
         TabIndex        =   31
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "City:"
         Height          =   255
         Index           =   2
         Left            =   840
         TabIndex        =   30
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Address:"
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   29
         Top             =   1440
         Width           =   1335
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
      ScaleWidth      =   6780
      TabIndex        =   27
      Top             =   5955
      Width           =   6780
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   300
         Left            =   4320
         TabIndex        =   21
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   300
         Left            =   3240
         TabIndex        =   20
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   300
         Left            =   2160
         TabIndex        =   19
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   300
         Left            =   1080
         TabIndex        =   18
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   300
         Left            =   0
         TabIndex        =   17
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
      ScaleWidth      =   6780
      TabIndex        =   0
      Top             =   6255
      Width           =   6780
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Picture         =   "frm_AR_Cust_Ship_To.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Picture         =   "frm_AR_Cust_Ship_To.frx":048C
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Picture         =   "frm_AR_Cust_Ship_To.frx":07CE
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Picture         =   "frm_AR_Cust_Ship_To.frx":0B10
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   26
         Top             =   0
         Width           =   3360
      End
   End
   Begin VB.Label lblTop 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Customer Data"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   42
      Top             =   120
      Width           =   5895
   End
End
Attribute VB_Name = "frm_AR_Cust_Ship_To"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ADOprimaryrs As ADODB.Recordset
Attribute ADOprimaryrs.VB_VarHelpID = -1
Dim db As ADODB.Connection
Dim TempCriteria As String

Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean

Dim NewLoad As Boolean

Private Sub bt_Cust_ShipTo_Click()
    Dim ghead As String
    Dim fhead As String
    
    ghead = "Ship To"
    fhead = "ID//Name//Address 1//Address 2"
    AllLookup.ToWhichRecord ADOprimaryrs, ghead, fhead
    'AllLookup.Show vbModal
End Sub

Private Sub chkFieldsq_Click(Index As Integer)
    'Doesn't work
  'If db Is Nothing Then   ' being a little bit paranoid
  '  Set db = New ADODB.Connection
  '  db.CursorLocation = adUseClient
  '  db.Open gblADOProvider
  'End If
  
  'to guarantee consistency in the database, only one default shipping for one _
    customer
  'Dim SQL As String
  'Dim SQLW As String
  'SQL = "select [AR SHIP ID],[AR SHIP Name],[AR SHIP Address 1],[AR SHIP Address 2]," & _
    "[AR SHIP Customer ID],[AR SHIP Default],[AR SHIP City],[AR SHIP State]," & _
    "[AR SHIP Postal],[AR SHIP Country],[AR SHIP Contact],[AR SHIP Phone]," & _
    "[AR SHIP Phone Ext],[AR SHIP Fax],[AR SHIP Notes] from [AR Ship To]"
  'SQLW = " where [AR SHIP Customer ID] = '" & frm_AR_Customer.txtFields(0).Text & "'"

  'Dim uchkrs As ADODB.Recordset
  'Dim i As Integer
  
  'Set uchkrs = New ADODB.Recordset
  'uchkrs.Open SQL & SQLW, db, adOpenDynamic, adLockOptimistic
  'uchkrs.Filter = "[AR SHIP Default] = TRUE"
  
  'For i = 1 To uchkrs.RecordCount
  '  uchkrs.Fields("AR SHIP Default") = False
  'Next
  
  'updates and closes recordset
  'uchkrs.UpdateBatch adAffectAll
  'While (uchkrs.State <> adStateExecuting)
  '  uchkrs.Close
  '  Set uchkrs = Nothing
  'Wend
  
  'Me.chkFields(15).Value = True
End Sub


Private Sub chkFields_LostFocus()
Dim sID As String
If NewLoad = True Or mbAddNewFlag = True Then Exit Sub
    If ADOprimaryrs.RecordCount > 1 Then
        mvBookMark = ADOprimaryrs.Bookmark ' = txtFields(2).Text
        ADOprimaryrs.MoveFirst
        Do While Not ADOprimaryrs.EOF
            ADOprimaryrs![AR SHIP Default] = False
            ADOprimaryrs.MoveNext
        Loop
        ADOprimaryrs.Bookmark = mvBookMark
        'ADOprimaryrs.Find "[AR SHIP ID]='" & sID & "'"
        ADOprimaryrs![AR SHIP Default] = True
        ADOprimaryrs.Update
    End If

End Sub

Private Sub Form_Load()
  Me.Width = 6900
  Me.Height = 6960
'On Error GoTo FormErr
    Set db = New ADODB.Connection
    db.CursorLocation = adUseClient
    db.Open gblADOProvider
  
  Dim sql As String
  Dim SQLW As String
  sql = "select [AR SHIP ID],[AR SHIP Name],[AR SHIP Address 1],[AR SHIP Address 2]," & _
    "[AR SHIP Customer ID],[AR SHIP Default],[AR SHIP City],[AR SHIP State]," & _
    "[AR SHIP Postal],[AR SHIP Country],[AR SHIP Contact],[AR SHIP Phone]," & _
    "[AR SHIP Phone Ext],[AR SHIP Fax],[AR SHIP Notes] from [AR Ship To]"
  SQLW = " where [AR SHIP Customer ID] = '" & TempCriteria & "'"
  
  Set ADOprimaryrs = New ADODB.Recordset
  
  ADOprimaryrs.Open sql & SQLW, db, adOpenKeyset, adLockOptimistic, adCmdText
  
  If ADOprimaryrs.RecordCount = 0 Then
    cmdUpdate.Enabled = False
    cmdDelete.Enabled = False
    cmdRefresh.Enabled = False
    bt_Cust_ShipTo.Enabled = False
    txtfields(2).Enabled = False
  End If
  
  Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtfields
    Set oText.DataSource = ADOprimaryrs
    If oText.DataField <> "" Then
        If ADOprimaryrs("" & oText.DataField & "").Type = 202 Then oText.MaxLength = ADOprimaryrs("" & oText.DataField & "").DefinedSize
    End If
  Next

  'Dim oCheck As CheckBox
    'Bind the Check boxes to the data provider
  'For Each oCheck In Me.chkFields
  '  Set oCheck.DataSource = ADOprimaryrs
  'Next

  If CheckNewDB(ADOprimaryrs, "Ship To Setup") = True Then
    cmdAdd_Click
    txtfields(0).Text = TempCriteria
    txtfields(2).Text = TempCriteria
  End If

  Set chkFields.DataSource = ADOprimaryrs

  'disable Customer ID textbox to avoid invalid enteries
  Me.txtfields(0).Enabled = False
  
  GetTextColor Me
  NewLoad = False
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
  
  Me.Width = 6900
  Me.Height = 6960
  
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
    'updates the checklist Projects
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
  ShowStatus False
  Set frm_AR_Cust_Ship_To = Nothing
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
On Error GoTo AddErr
  Dim flg As Boolean
  
  With ADOprimaryrs
    If cmdAdd.Caption = "&Add" Then
        If Not (.BOF And .EOF) Then
            mvBookMark = .Bookmark
        End If
        .AddNew
        lblStatus.Caption = "Add record"
        txtfields(2).Enabled = True
        cmdAdd.Caption = "&Cancel"
        cmdUpdate.Enabled = True
        mbAddNewFlag = True
        flg = False
    Else
        .CancelUpdate
        If .RecordCount > 0 Then
            If mvBookMark > 0 Then
                .Bookmark = mvBookMark
            Else
                .MoveLast
            End If
        End If
        cmdAdd.Caption = "&Add"
        txtfields(2).Enabled = False
        mbAddNewFlag = False
        flg = True
    End If
    
    'enable/disable the proper buttons
    cmdDelete.Enabled = flg
    cmdRefresh.Enabled = flg
    bt_Cust_ShipTo.Enabled = flg
    
    mbAddNewFlag = True
  End With
  
  'Automatically assignes the Customer ID
  Me.txtfields(0).Text = TempCriteria

  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
  'On Error GoTo DeleteErr
  With ADOprimaryrs
    If .RecordCount = 0 Then Exit Sub   ' no records maa...
    If .EditMode = False Then
        .Delete
        .MoveNext
        If .RecordCount = 0 Then
            cmdUpdate.Enabled = False
            cmdDelete.Enabled = False
            cmdRefresh.Enabled = False
            bt_Cust_ShipTo.Enabled = False
            txtfields(2).Enabled = False
            .Requery
            Exit Sub
        ElseIf .EOF Then .MoveLast
        End If
    Else
        MsgBox "Must update or refresh before deleting", vbCritical, _
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
            .Requery                'cancels any modification in form
            .Bookmark = mvBookMark
        End If
    End With
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub

Private Sub cmdCancel_Click()
  'On Error Resume Next

  mbEditFlag = False
  mbAddNewFlag = False
  ADOprimaryrs.CancelUpdate
  If mvBookMark > 0 Then
    ADOprimaryrs.Bookmark = mvBookMark
  Else
    ADOprimaryrs.MoveFirst
  End If
  mbDataChanged = False

End Sub

Private Sub cmdUpdate_Click()
  'On Error GoTo UpdateErr

    With ADOprimaryrs
        If .RecordCount = 0 Then Exit Sub 'no records to update
        If Trim(txtfields(0).Text) <> "" Then
        Dim oTxt As TextBox
          For Each oTxt In Me.txtfields
            If oTxt.Text = "" And oTxt.DataField <> "" Then
              If ADOprimaryrs("" & oTxt.DataField & "").Type = 203 Or ADOprimaryrs("" & oTxt.DataField & "").Type = 202 Then oTxt.Text = " "
            End If
          Next
        End If
        If .RecordCount = 0 Then Exit Sub ' no records to update
        .Update
        If mbAddNewFlag Then    'gets default database values
            .Requery
            .MoveLast
            mbAddNewFlag = False
        End If
        
        'reenable necessary buttons
        cmdAdd.Caption = "&Add"
        txtfields(2).Enabled = False
        cmdDelete.Enabled = True
        cmdRefresh.Enabled = True
        bt_Cust_ShipTo.Enabled = True
    End With
  
  mbEditFlag = False
  mbDataChanged = False

  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdFirst_Click()
  'On Error GoTo GoFirstError

  ADOprimaryrs.MoveFirst
'  mbDataChanged = False

  Exit Sub

GoFirstError:
  MsgBox Err.Description
End Sub

Private Sub cmdLast_Click()
  'On Error GoTo GoLastError

  ADOprimaryrs.MoveLast
'  mbDataChanged = False

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
'  mbDataChanged = False

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
'  mbDataChanged = False

  Exit Sub

GoPrevError:
  MsgBox Err.Description
End Sub

Private Sub SetButtons(bVal As Boolean)
  cmdAdd.Visible = bVal
  cmdUpdate.Visible = bVal
  cmdDelete.Visible = bVal
  cmdClose.Visible = bVal
  cmdRefresh.Visible = bVal
  cmdNext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdLast.Enabled = bVal
  cmdPrevious.Enabled = bVal
End Sub

Public Sub CallByUserShip(shipToID As String, Optional ARCustCall As Boolean)
    NewLoad = True
    TempCriteria = shipToID
    Me.Show
If ARCustCall = False Then
    If mbAddNewFlag = False Then
        cmdAdd_Click
    End If
End If
End Sub

