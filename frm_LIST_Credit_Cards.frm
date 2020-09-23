VERSION 5.00
Begin VB.Form frm_LIST_Credit_Cards 
   Caption         =   "Maintain Company Credit Cards"
   ClientHeight    =   3795
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   5400
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3795
   ScaleWidth      =   5400
   Begin VB.PictureBox picStatBox 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   5400
      TabIndex        =   24
      Top             =   3495
      Width           =   5400
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Picture         =   "frm_LIST_Credit_Cards.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Picture         =   "frm_LIST_Credit_Cards.frx":0342
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Picture         =   "frm_LIST_Credit_Cards.frx":0684
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Picture         =   "frm_LIST_Credit_Cards.frx":09C6
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   25
         Top             =   0
         Width           =   3360
      End
   End
   Begin VB.Frame frPrimary 
      Height          =   2655
      Left            =   0
      TabIndex        =   17
      Top             =   480
      Width           =   5295
      Begin VB.CommandButton cmdDate 
         Height          =   285
         Left            =   4695
         Picture         =   "frm_LIST_Credit_Cards.frx":0D08
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "LIST CC Name"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   0
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox txtFields 
         DataField       =   "LIST CC Exp Date"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "MM/dd/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   3
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1680
         Width           =   2775
      End
      Begin VB.TextBox txtFields 
         DataField       =   "LIST CC Cardholder"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   1
         Left            =   1920
         TabIndex        =   2
         Top             =   960
         Width           =   3135
      End
      Begin VB.TextBox txtFields 
         DataField       =   "LIST CC Card No"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   2
         Left            =   1920
         TabIndex        =   3
         Top             =   1320
         Width           =   3135
      End
      Begin VB.TextBox txtFields 
         DataField       =   "LIST Account Number"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   4
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   2040
         Width           =   2775
      End
      Begin VB.CommandButton btCCardType 
         Height          =   285
         Left            =   4680
         Picture         =   "frm_LIST_Credit_Cards.frx":1012
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton btCCardAcc 
         Height          =   285
         Left            =   4680
         Picture         =   "frm_LIST_Credit_Cards.frx":115C
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2040
         Width           =   375
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Credit Card Type"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   22
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Expiration Date"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   21
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Name on Credit Card"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   20
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Credit Card Number"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   19
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "GL Account Number"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   18
         Top             =   2040
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
      ScaleWidth      =   5400
      TabIndex        =   16
      Top             =   3195
      Width           =   5400
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
   Begin VB.Label lblTop 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Employee Types"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   23
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frm_LIST_Credit_Cards"
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

Private Sub btCCardAcc_Click()
    Dim SQLstatement As String
    Dim No As Integer
    Dim ghead As String
    Dim fhead As String

    No = 2
    SQLstatement = "select [GL COA Account No], [GL COA Account Name]" & _
                    "from [GL Chart Of Accounts]"
    ghead = "Account Description"
    fhead = "Account No//Description"
    
    AllLookup.GetWhichTable No, SQLstatement, ghead, fhead, db
    txtFields(4).SetFocus   'trigger event adFirstChange
    'AllLookup.Show vbModal
End Sub

Private Sub btCCardType_Click()
    Dim ghead As String
    Dim fhead As String
        
    ghead = "Credit Card Type"
    fhead = "Type//Card Holder//Card Number"
    AllLookup.ToWhichRecord ADOprimaryrs, ghead, fhead
    'AllLookup.Show vbModal
End Sub

Private Sub cmdDate_Click()
    Menu_Calendar.WhoCallMe True, 1565
    'Menu_Calendar.Show vbModal
    txtFields(3).SetFocus
End Sub

Private Sub Form_Load()
ShowStatus True
On Error GoTo FormErr
    Set db = New ADODB.Connection
    db.CursorLocation = adUseClient
    db.Open gblADOProvider

  Dim sql As String
  sql = "select [LIST CC Name],[LIST CC Cardholder],[LIST CC Card No]," & _
    "[LIST CC Exp Date],[LIST Account Number] from [LIST Credit Cards]"
    
  Set ADOprimaryrs = New ADODB.Recordset
  ADOprimaryrs.Open sql, db, adOpenStatic, adLockOptimistic

  Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtFields
  '202 - text 203 - memo 7 - date 5 - number
    Set oText.DataSource = ADOprimaryrs
    If ADOprimaryrs("" & oText.DataField & "").Type = 202 Then oText.MaxLength = ADOprimaryrs("" & oText.DataField & "").DefinedSize
  Next

  If CheckNewDB(ADOprimaryrs, "Credit Card") = True Then
    cmdAdd_Click
  End If

  Me.Width = 5520
  Me.Height = 4200

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
  
  Me.Width = 5520
  Me.Height = 4200
  
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
  'updates the checklist Credit Cards Used
  ShowStatus True
      EndLoad db, ADOprimaryrs, "Credit Cards Used"
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
  Set frm_LIST_Credit_Cards = Nothing
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
            txtFields(0).Locked = False
            'txtfields(0).SetFocus
        Else
            mbAddNewFlag = False
            txtFields(0).Locked = True
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
        btCCardAcc.Enabled = True
        btCCardType.Enabled = Not mbAddNewFlag
        cmdDelete.Enabled = Not mbAddNewFlag
        cmdRefresh.Enabled = Not mbAddNewFlag
GetTextColor Me
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
        If Trim(txtFields(0).Text) <> "" Then
        Dim oTxt As TextBox
          For Each oTxt In Me.txtFields
            If oTxt.Text = "" Then
              If ADOprimaryrs("" & oTxt.DataField & "").Type = 203 Or ADOprimaryrs("" & oTxt.DataField & "").Type = 202 Then oTxt.Text = " "
            End If
          Next
        End If
        .Update
        If mbAddNewFlag Then
            .Requery            ' gets default value assigned by database
            .MoveLast
            mbAddNewFlag = False
        End If
    End With
    
    'reenable the neccessary buttons
    cmdAdd.Caption = "&Add"
    btCCardAcc.Enabled = True
    btCCardType.Enabled = True
    cmdUpdate.Enabled = True
    cmdDelete.Enabled = True
    cmdRefresh.Enabled = True
GetTextColor Me
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
  cmdUpdate.Visible = bVal
  cmdDelete.Visible = bVal
  cmdClose.Visible = bVal
  cmdRefresh.Visible = bVal
  cmdNext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdLast.Enabled = bVal
  cmdPrevious.Enabled = bVal
End Sub
