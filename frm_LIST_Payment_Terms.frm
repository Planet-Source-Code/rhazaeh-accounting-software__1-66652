VERSION 5.00
Begin VB.Form frm_LIST_Payment_Terms 
   Caption         =   "Maintain Payment Terms"
   ClientHeight    =   4035
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   6180
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4035
   ScaleWidth      =   6180
   Begin VB.Frame frPrimary 
      Height          =   2895
      Left            =   0
      TabIndex        =   19
      Top             =   480
      Width           =   6135
      Begin VB.TextBox txtFields 
         DataField       =   "LIST PAY Net Days"
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
         Index           =   5
         Left            =   2400
         TabIndex        =   3
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox txtFields 
         DataField       =   "LIST PAY ID"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   4
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtFields 
         DataField       =   "LIST PAY Expire Day"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "MM/dd/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   3
         Left            =   2400
         TabIndex        =   6
         Top             =   2280
         Width           =   1935
      End
      Begin VB.TextBox txtFields 
         DataField       =   "LIST PAY Discount Days"
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
         Index           =   2
         Left            =   2400
         TabIndex        =   5
         Top             =   1920
         Width           =   1935
      End
      Begin VB.TextBox txtFields 
         DataField       =   "LIST PAY Discount"
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
         Index           =   1
         Left            =   2400
         TabIndex        =   4
         Top             =   1560
         Width           =   1935
      End
      Begin VB.TextBox txtFields 
         DataField       =   "LIST PAY Description"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   0
         Left            =   2400
         TabIndex        =   2
         Top             =   840
         Width           =   3375
      End
      Begin VB.CommandButton btTermsID 
         Height          =   285
         Left            =   3960
         Picture         =   "frm_LIST_Payment_Terms.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   480
         Width           =   375
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Net Days Due"
         Height          =   255
         Index           =   5
         Left            =   480
         TabIndex        =   25
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Terms ID"
         Height          =   255
         Index           =   4
         Left            =   480
         TabIndex        =   24
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Discount Expiration Days"
         Height          =   255
         Index           =   3
         Left            =   480
         TabIndex        =   23
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Discount Days"
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   22
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Terms Discount"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   21
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Terms Description"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   20
         Top             =   840
         Width           =   1815
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
      ScaleWidth      =   6180
      TabIndex        =   18
      Top             =   3435
      Width           =   6180
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   300
         Left            =   4320
         TabIndex        =   12
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   300
         Left            =   3240
         TabIndex        =   11
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   300
         Left            =   2160
         TabIndex        =   10
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   300
         Left            =   1080
         TabIndex        =   9
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   300
         Left            =   0
         TabIndex        =   8
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
      ScaleWidth      =   6180
      TabIndex        =   7
      Top             =   3735
      Width           =   6180
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Picture         =   "frm_LIST_Payment_Terms.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Picture         =   "frm_LIST_Payment_Terms.frx":048C
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Picture         =   "frm_LIST_Payment_Terms.frx":07CE
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Picture         =   "frm_LIST_Payment_Terms.frx":0B10
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   17
         Top             =   0
         Width           =   3360
      End
   End
   Begin VB.Label lblTop 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Project Types"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   26
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frm_LIST_Payment_Terms"
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

Private Sub btTermsID_Click()
    Dim ghead As String
    Dim fhead As String
        
    ghead = "Payment Terms"
    fhead = "ID//Description"
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
  sql = "select [LIST PAY ID],[LIST PAY Description],[LIST PAY Net Days]," & _
    "[LIST PAY Discount],[LIST PAY Discount Days],[LIST PAY Expire Day] " & _
    "from [LIST Payment Terms]"
    
  Set ADOprimaryrs = New ADODB.Recordset
  ADOprimaryrs.Open sql, db, adOpenStatic, adLockOptimistic
  
  'disable Term ID to avoid invalid entry
  txtfields(4).Locked = True

  Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtfields
    Set oText.DataSource = ADOprimaryrs
    If ADOprimaryrs("" & oText.DataField & "").Type = 202 Then oText.MaxLength = ADOprimaryrs("" & oText.DataField & "").DefinedSize
  Next
  
  If CheckNewDB(ADOprimaryrs, "Payment Terms") = True Then
    cmdAdd_Click
  End If

  Me.Width = 6300
  Me.Height = 4440

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
  
  Me.Width = 6300
  Me.Height = 4440
  
SkipResize:
  frPrimary.Left = (Me.ScaleWidth - frPrimary.Width) / 2
  lblTop.Left = frPrimary.Left
  lblTop.Width = frPrimary.Width
  frPrimary.Top = (Me.ScaleHeight - frPrimary.Height - picButtons.Height - picStatBox.Height) / 2 + 230
  
  lblStatus.Width = Me.Width - 1500
  cmdNext.Left = lblStatus.Width + 700
  cmdLast.Left = cmdNext.Left + 340
End Sub

Public Sub CallByUserPay(TermsID As String)
    Me.Show
    If mbAddNewFlag = False Then
        cmdAdd_Click
        txtfields(0) = TermsID
    Else
        txtfields(0) = TermsID
    End If
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
      EndLoad db, ADOprimaryrs, "Payment Terms"
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
Set frm_LIST_Payment_Terms = Nothing
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
            txtfields(0).Locked = False
            txtfields(0).SetFocus
        Else
            mbAddNewFlag = False
            .CancelUpdate
            If .RecordCount > 0 Then
                If mvBookMark > 0 Then
                    .Bookmark = mvBookMark
                Else
                    .MoveLast
                End If
            End If
        End If
        
        'set the buttons accordingly
        If mbAddNewFlag Then
            cmdAdd.Caption = "&Cancel"
        Else
            cmdAdd.Caption = "&Add"
        End If
        cmdUpdate.Enabled = True
        btTermsID.Enabled = Not mbAddNewFlag
        cmdDelete.Enabled = Not mbAddNewFlag
        cmdRefresh.Enabled = Not mbAddNewFlag
    End With

  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
  'On Error GoTo DeleteErr
  With ADOprimaryrs
    If .RecordCount = 0 Then Exit Sub ' no records maa....
    If .EditMode = False Then
        .Delete
        .MoveNext
        If .RecordCount = 0 Then
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
        MsgBox "Must update or refresh record before deleting.", vbCritical, "Delete Error."
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
        If Not (.BOF Or .EOF) Then
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
  'On Error GoTo UpdateErr

    With ADOprimaryrs
        Dim oTxt As TextBox
          For Each oTxt In Me.txtfields
            If oTxt.Text = "" And oTxt.Index <> 4 Then
              'If ADOprimaryrs("" & oTxt.DataField & "").Type = 203 Or ADOprimaryrs("" & oTxt.DataField & "").Type = 202 Then oTxt.Text = " "
              MsgBox lblLabels(oTxt.Index) & " cannot be left empty", vbInformation, "Information"
              Exit Sub
            End If
          Next
        .Update
        If mbAddNewFlag Then
            .Requery
            .MoveLast
            mbAddNewFlag = False
        End If
    End With
    
    'reenable the neccessary buttons
    cmdAdd.Caption = "&Add"
    btTermsID.Enabled = True
    cmdUpdate.Enabled = True
    cmdDelete.Enabled = True
    cmdRefresh.Enabled = True
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
Select Case Index
Case 1, 2, 3, 5
    keyResponse = CtrlValidate(KeyAscii, "0123456789.")
    If keyResponse = True Then
    Else
       KeyAscii = 0
    End If
End Select
End Sub

Private Sub txtFields_LostFocus(Index As Integer)
Select Case Index
Case 1, 2, 3, 5
    txtfields(Index) = Val(txtfields(Index))
End Select
End Sub
