VERSION 5.00
Begin VB.Form frm_SYS_Setup_Company 
   Caption         =   "Maintain Company Information"
   ClientHeight    =   5685
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   9765
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5685
   ScaleWidth      =   9765
   Begin VB.Frame frPrimary 
      Height          =   4815
      Left            =   0
      TabIndex        =   20
      Top             =   480
      Width           =   9735
      Begin VB.CheckBox chkFields 
         Alignment       =   1  'Right Justify
         Caption         =   "Use Passwords"
         DataField       =   "SYS COM Use Passwords YN"
         DataSource      =   "adoprimaryrs"
         Enabled         =   0   'False
         Height          =   195
         Index           =   2
         Left            =   3840
         TabIndex        =   35
         Top             =   3720
         Width           =   2055
      End
      Begin VB.PictureBox picSYSLOGO 
         BorderStyle     =   0  'None
         DataField       =   "SYS COM Logo"
         Height          =   1815
         Left            =   6240
         ScaleHeight     =   1815
         ScaleWidth      =   3375
         TabIndex        =   34
         Top             =   240
         Width           =   3375
      End
      Begin VB.CommandButton cmdPic 
         Height          =   285
         Left            =   5520
         Picture         =   "frm_SYS_Setup_Company.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Load Company Logo"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "SYS COM Web"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   10
         Left            =   2040
         TabIndex        =   31
         Top             =   3960
         Width           =   3855
      End
      Begin VB.CheckBox chkFields 
         Alignment       =   1  'Right Justify
         Caption         =   "Use Departments"
         DataField       =   "SYS COM Use Departments YN"
         DataSource      =   "adoPrimaryRS"
         Enabled         =   0   'False
         Height          =   255
         Index           =   13
         Left            =   3840
         TabIndex        =   11
         Top             =   1800
         Width           =   2055
      End
      Begin VB.TextBox txtFields 
         DataField       =   "SYS COM Notes"
         DataSource      =   "adoPrimaryRS"
         Height          =   2055
         Index           =   15
         Left            =   6240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   2400
         Width           =   3375
      End
      Begin VB.CheckBox chkFields 
         Alignment       =   1  'Right Justify
         Caption         =   "Show Tips"
         DataField       =   "SYS COM Show Tips YN"
         DataSource      =   "adoPrimaryRS"
         Enabled         =   0   'False
         Height          =   255
         Index           =   12
         Left            =   3840
         TabIndex        =   12
         Top             =   2280
         Width           =   2055
      End
      Begin VB.TextBox txtFields 
         DataField       =   "SYS COM Fax"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   9
         Left            =   2040
         TabIndex        =   10
         Top             =   3600
         Width           =   1575
      End
      Begin VB.TextBox txtFields 
         DataField       =   "SYS COM Phone"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   8
         Left            =   2040
         TabIndex        =   9
         Top             =   3240
         Width           =   1575
      End
      Begin VB.TextBox txtFields 
         DataField       =   "SYS COM Country"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   7
         Left            =   2040
         TabIndex        =   8
         Top             =   2880
         Width           =   1575
      End
      Begin VB.TextBox txtFields 
         DataField       =   "SYS COM Postal"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   6
         Left            =   2040
         TabIndex        =   7
         Top             =   2520
         Width           =   1575
      End
      Begin VB.TextBox txtFields 
         DataField       =   "SYS COM State"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   5
         Left            =   2040
         TabIndex        =   6
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox txtFields 
         DataField       =   "SYS COM City"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   4
         Left            =   2040
         TabIndex        =   5
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox txtFields 
         DataField       =   "SYS COM Address 2"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   3
         Left            =   2040
         TabIndex        =   4
         Top             =   1440
         Width           =   3855
      End
      Begin VB.TextBox txtFields 
         DataField       =   "SYS COM Address 1"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   2
         Left            =   2040
         TabIndex        =   3
         Top             =   1080
         Width           =   3855
      End
      Begin VB.TextBox txtFields 
         DataField       =   "SYS COM Company Name"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   1
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   2
         Top             =   720
         Width           =   3855
      End
      Begin VB.TextBox txtFields 
         DataField       =   "SYS COM ID"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   0
         Left            =   2040
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
      Begin VB.CheckBox chkFields 
         Alignment       =   1  'Right Justify
         Caption         =   "Department"
         DataField       =   "SYS COM Department YN"
         DataSource      =   "adoprimaryrs"
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   3840
         TabIndex        =   13
         Top             =   3240
         Width           =   2055
      End
      Begin VB.CheckBox chkFields 
         Alignment       =   1  'Right Justify
         Caption         =   "Register to tbs.com.my"
         DataField       =   "SYS COM Registration"
         DataSource      =   "adoprimaryrs"
         Enabled         =   0   'False
         Height          =   195
         Index           =   1
         Left            =   3840
         TabIndex        =   14
         Top             =   2760
         Width           =   2055
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Company Web Page:  "
         Height          =   255
         Index           =   10
         Left            =   480
         TabIndex        =   32
         Top             =   3960
         Width           =   1575
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Business Description"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   15
         Left            =   6240
         TabIndex        =   30
         Top             =   2160
         Width           =   3375
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Fax Number:  "
         Height          =   255
         Index           =   9
         Left            =   480
         TabIndex        =   29
         Top             =   3600
         Width           =   1575
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Telephone Number:  "
         Height          =   255
         Index           =   8
         Left            =   480
         TabIndex        =   28
         Top             =   3240
         Width           =   1575
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Country:  "
         Height          =   255
         Index           =   7
         Left            =   480
         TabIndex        =   27
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Zip Code:  "
         Height          =   255
         Index           =   6
         Left            =   480
         TabIndex        =   26
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "State:  "
         Height          =   255
         Index           =   5
         Left            =   480
         TabIndex        =   25
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "City:  "
         Height          =   255
         Index           =   4
         Left            =   480
         TabIndex        =   24
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Address:  "
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   23
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Company Name:  "
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   22
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Company ID:  "
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   21
         Top             =   360
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
      ScaleWidth      =   9765
      TabIndex        =   0
      Top             =   5385
      Width           =   9765
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   300
         Left            =   2160
         TabIndex        =   18
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   300
         Left            =   1080
         TabIndex        =   17
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   300
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.Label lbltop 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Company Information"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   30
      TabIndex        =   19
      Top             =   120
      Width           =   9700
   End
End
Attribute VB_Name = "frm_SYS_Setup_Company"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ADOprimaryrs As ADODB.Recordset
Attribute ADOprimaryrs.VB_VarHelpID = -1
Dim db As ADODB.Connection
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean 'SYS COM Registration
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean
' Note:
' The recordset/table for this form should only contain one row with information
' for a specific company. No new records should be added to this table.

Private Sub cmdPic_Click()
   Dim bytData() As Byte
   'Dim sFile As String
    
    sFile = ""
    With fMainForm.dlgCommonDialog
        .DialogTitle = "Load Picture For ToolBar"
        .CancelError = False
        .Flags = cdlOFNPathMustExist + cdlOFNOverwritePrompt
        .Filter = "All Files (*.BMP;*.Gif;*.JPG)|*.BMP;*.Gif;*.JPG"
        .ShowOpen
        If Len(.FileName) = 0 Then
            ShowStatus False
            Exit Sub
        End If
        If .FileName = "" Then Exit Sub
        
        'Open the picture file
        Open .FileName For Binary As #1
        ReDim bytData(FileLen(.FileName))
        
    End With
    'If sFile <> "" Then
        'txtLocked(6).Text = sFile
        'txtLocked(6).SetFocus
        'picSYSLOGO.Picture = LoadPicture(sFile)
        'ADOprimaryrs![SYS COM Logo] = sFile
        'fMainForm.CoolBar1.Refresh
        'fMainForm.tbToolBar1.RestoreToolbar
        'cmdApply.Enabled = True
    'End If
    'Read the data and close the file
    Get #1, , bytData
    Close #1
    
    'Add the record
    With ADOprimaryrs
        .Fields("SYS COM Logo").AppendChunk bytData
        .Update
    End With
    If ADOprimaryrs.RecordCount = 1 Then ADOprimaryrs.MoveFirst
End Sub

Private Sub Form_Load()
ShowStatus True
'On Error GoTo FormErr
  Set db = New ADODB.Connection
  db.CursorLocation = adUseClient
  db.Open gblADOProvider

  Set ADOprimaryrs = New ADODB.Recordset
  ADOprimaryrs.Open "select [SYS COM ID],[SYS COM Company Name],[SYS COM Address 1]," & _
    "[SYS COM Address 2],[SYS COM City],[SYS COM State],[SYS COM Postal]," & _
    "[SYS COM Country],[SYS COM Phone],[SYS COM Fax],[SYS COM Department YN]," & _
    "[SYS COM Show Tips YN],[SYS COM Use Departments YN],[SYS COM Web],[SYS COM Logo]," & _
    "[SYS COM Registration],[SYS COM Use Passwords YN],[SYS COM Notes] from [SYS Company] " & _
    "Order by [SYS COM ID]", db, adOpenStatic, adLockOptimistic, adCmdText

  Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtFields
    Set oText.DataSource = ADOprimaryrs
    If ADOprimaryrs("" & oText.DataField & "").Type = 202 Then oText.MaxLength = ADOprimaryrs("" & oText.DataField & "").DefinedSize
  Next

  Dim oCheck As CheckBox
    'Bind the Check boxes to the data provider
  For Each oCheck In Me.chkFields
    Set oCheck.DataSource = ADOprimaryrs
  Next
  
  Set picSYSLOGO.DataSource = ADOprimaryrs
  
  If CheckNewDB(ADOprimaryrs, "Company") = True Then
    ADOprimaryrs.AddNew
  End If
  
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
  
  Me.Width = 9855
  Me.Height = 6090
  
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
    'updates the checklist Company
  ShowStatus True
      EndLoad db, ADOprimaryrs, "Company"
      If ADOprimaryrs.RecordCount > 0 Then
        If ADOprimaryrs.EditMode <> 0 Then
          ADOprimaryrs.CancelUpdate
        End If
      End If
      ADOprimaryrs.Close
      Set ADOprimaryrs = Nothing
      db.Close
      Set db = Nothing
      Set frm_SYS_Setup_Company = Nothing
  ShowStatus False
Exit Sub
FormErr:
  MsgBox Err.Description
  ShowStatus False
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
Dim oTxt As TextBox

  If Trim(txtFields(0).Text) <> "" Then
    For Each oTxt In Me.txtFields
      If oTxt.Text = "" Then
          oTxt.Text = " "
      End If
    Next
  End If
  
  With ADOprimaryrs
  .Update
'  .Requery
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

Private Sub txtfields_GotFocus(Index As Integer)
    TxtGotFocus txtFields(Index)
End Sub
