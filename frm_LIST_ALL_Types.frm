VERSION 5.00
Begin VB.Form frm_LIST_ALL_Types 
   Caption         =   "Maintain Customer Types"
   ClientHeight    =   4755
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   6945
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4755
   ScaleWidth      =   6945
   Begin VB.Frame frPrimary 
      Height          =   3615
      Left            =   0
      TabIndex        =   12
      Top             =   480
      Width           =   6855
      Begin VB.TextBox txtFields 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """$""#,##0.00;(""$""#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   4680
         TabIndex        =   28
         Top             =   840
         Width           =   2055
      End
      Begin VB.CommandButton cmdUpdatedua 
         Height          =   280
         Left            =   6360
         Picture         =   "frm_LIST_ALL_Types.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Refresh Payment Terms"
         Top             =   2160
         Width           =   375
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   4680
         TabIndex        =   25
         Text            =   "Combo1"
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Customer Type"
         DataSource      =   "adoPrimaryRS"
         Height          =   285
         Index           =   0
         Left            =   3360
         TabIndex        =   18
         Top             =   1560
         Width           =   3375
      End
      Begin VB.Frame frSelection 
         Height          =   3375
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   3015
         Begin VB.OptionButton optType 
            Caption         =   "GL Account Types"
            Height          =   255
            Index           =   7
            Left            =   360
            TabIndex        =   29
            Top             =   2880
            Width           =   2415
         End
         Begin VB.OptionButton optType 
            Caption         =   "Recurring Type"
            Height          =   255
            Index           =   5
            Left            =   360
            TabIndex        =   23
            Top             =   1080
            Width           =   2415
         End
         Begin VB.OptionButton optType 
            Caption         =   "Vendor Type"
            Height          =   255
            Index           =   4
            Left            =   360
            TabIndex        =   22
            Top             =   720
            Width           =   2415
         End
         Begin VB.OptionButton optType 
            Caption         =   "Item Type"
            Height          =   255
            Index           =   6
            Left            =   360
            TabIndex        =   21
            Top             =   1440
            Width           =   2415
         End
         Begin VB.OptionButton optType 
            Caption         =   "Payment Methods"
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   17
            Top             =   2160
            Width           =   2415
         End
         Begin VB.OptionButton optType 
            Caption         =   "Item Categories"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   16
            Top             =   1800
            Width           =   2415
         End
         Begin VB.OptionButton optType 
            Caption         =   "Customer Types"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   15
            Top             =   360
            Width           =   2415
         End
         Begin VB.OptionButton optType 
            Caption         =   "Shipping Methods"
            Height          =   255
            Index           =   3
            Left            =   360
            TabIndex        =   14
            Top             =   2520
            Width           =   2415
         End
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Shipping Charge:"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   3360
         TabIndex        =   27
         Top             =   840
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Payment Terms:"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   3360
         TabIndex        =   24
         Top             =   2160
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Customer Type:"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   3360
         TabIndex        =   19
         Top             =   1320
         Width           =   3375
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
      ScaleWidth      =   6945
      TabIndex        =   6
      Top             =   4155
      Width           =   6945
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   300
         Left            =   4320
         TabIndex        =   10
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   300
         Left            =   3240
         TabIndex        =   9
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   300
         Left            =   2160
         TabIndex        =   8
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   300
         Left            =   1080
         TabIndex        =   11
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
      ScaleWidth      =   6945
      TabIndex        =   0
      Top             =   4455
      Width           =   6945
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Picture         =   "frm_LIST_ALL_Types.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Picture         =   "frm_LIST_ALL_Types.frx":064C
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Picture         =   "frm_LIST_ALL_Types.frx":098E
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Picture         =   "frm_LIST_ALL_Types.frx":0CD0
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   5
         Top             =   0
         Width           =   3360
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
      TabIndex        =   20
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frm_LIST_ALL_Types"
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
Dim AllreadyOpen As Boolean

Private Sub OpenDB(SQLstatement As String)
  If AllreadyOpen = True Then
    ADOprimaryrs.CancelUpdate
    ADOprimaryrs.Close
  Else
    Set ADOprimaryrs = New ADODB.Recordset
    Me.Show
  End If
  AllreadyOpen = True

  ADOprimaryrs.Open SQLstatement, db, adOpenStatic, adLockOptimistic
  
  Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtfields
    'Set oText.DataSource = adoPrimaryRS
    If Trim(oText.DataField) <> "" Then
      If ADOprimaryrs("" & oText.DataField & "").Type = 202 Then oText.MaxLength = ADOprimaryrs("" & oText.DataField & "").DefinedSize
      Set oText.DataSource = ADOprimaryrs
    End If
  Next
  'MsgBox lbllabels(0).Caption
  If lblLabels(0).Caption = "Payment Methods:" Then
    Set Combo1.DataSource = ADOprimaryrs
  ElseIf lblLabels(0).Caption = "Shipping Methods:" Then
    Set txtfields(1).DataSource = ADOprimaryrs
  ElseIf lblLabels(0).Caption = "GL Account Types:" Then
    Set txtfields(1).DataSource = ADOprimaryrs
  Else
    Set Combo1.DataSource = Nothing
    Set txtfields(1).DataSource = Nothing
  End If
  mbDataChanged = False
End Sub

Private Sub cmdUpdatedua_Click()
    ComboInit Combo1, lblLabels(1), "SELECT [LIST PAY Description] FROM [LIST Payment Terms]", db
End Sub

Private Sub Form_Load()
ShowStatus True
    
    AllreadyOpen = False
    GetTextColor Me
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
  
  Me.Width = 7035
  Me.Height = 5160
  
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
Dim ADOsecondaryRS As ADODB.Recordset
On Error GoTo FormErr

  'updates the checklist Credit Cards Used
  ShowStatus True
        
      Set ADOsecondaryRS = New ADODB.Recordset
      ADOsecondaryRS.Open "select [LIST SHIP Method] from [LIST Shipping Methods]", db, adOpenKeyset, adLockOptimistic, adCmdText
      EndLoad db, ADOprimaryrs, "Shipping Methods", 1
      ADOsecondaryRS.Close
      Set ADOsecondaryRS = Nothing
    
      Set ADOsecondaryRS = New ADODB.Recordset
      ADOsecondaryRS.Open "select Category from [LIST Item Categories]", db, adOpenKeyset, adLockOptimistic, adCmdText
      EndLoad db, ADOprimaryrs, "Item Categories"
      ADOsecondaryRS.Close
      Set ADOsecondaryRS = Nothing
      
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
  Set frm_LIST_ALL_Types = Nothing
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
        .Update
        If mbAddNewFlag Then
            .Requery
            .MoveLast
            mbAddNewFlag = False
        End If
    End With
    
    'reenable the neccessary buttons
    cmdAdd.Caption = "&Add"
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
  cmdUpdate.Visible = bVal
  cmdDelete.Visible = bVal
  cmdClose.Visible = bVal
  cmdRefresh.Visible = bVal
  cmdNext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdLast.Enabled = bVal
  cmdPrevious.Enabled = bVal
End Sub

Private Sub optType_Click(Index As Integer)
'clear all datasource

If mbAddNewFlag Then
    cmdAdd_Click
End If

  If AllreadyOpen = True Then
    ADOprimaryrs.CancelUpdate
    ADOprimaryrs.Close
    AllreadyOpen = False
  End If
  
  For Each oText In Me.txtfields
    'Set oText.DataSource = adoPrimaryRS
     Set oText.DataSource = Nothing
  Next
  Set Combo1.DataSource = Nothing
  
    Combo1.DataField = ""
    Combo1.Visible = False
    lblLabels(1).Visible = False
    cmdUpdatedua.Visible = False
    txtfields(1).Visible = False
    lblLabels(2).Visible = False
    txtfields(1).DataField = ""
    lblLabels(2).Top = lblLabels(1).Top
    txtfields(1).Top = lblLabels(2).Top
      
lblLabels(0).Caption = optType(Index).Caption & ":"
Me.Caption = optType(Index).Caption

Select Case Index
Case 0
    txtfields(0).DataField = "Customer Type"
    OpenDB "select [Customer Type] from [LIST Customer Types]"
Case 1
    txtfields(0).DataField = "Category"
    OpenDB "select Category from [LIST Item Categories]"
Case 2
    txtfields(0).DataField = "LIST PAY Method"
    Combo1.DataField = "Payment Terms"
    Combo1.Visible = True
    lblLabels(1).Visible = True
    cmdUpdatedua.Visible = True
    ComboInit Combo1, lblLabels(1), "SELECT [LIST PAY Description] FROM [LIST Payment Terms]", db
    OpenDB "select [Payment Terms],[LIST PAY Method] from [LIST Payment Methods]"
Case 3 '
    txtfields(1).Visible = True
    lblLabels(2).Visible = True
    lblLabels(2).Caption = "Shipping Charge:"
    txtfields(1).DataField = "LIST SHIP Charge"
    txtfields(0).DataField = "LIST SHIP Method"
    OpenDB "select [id],[LIST SHIP Charge],[LIST SHIP Method] from [LIST Shipping Methods]"
Case 4
    txtfields(0).DataField = "LIST VENDOR Types"
    OpenDB "select [LIST VENDOR Types] from [LIST Vendor Types]"
Case 5 '
    txtfields(0).DataField = "RECURR TYPE"
    OpenDB "select [RECURR TYPE] from [RECUR_TYPE]"
Case 6
    txtfields(0).DataField = "Type"
    OpenDB "select Type from [LIST Item Types]"
Case 7
    txtfields(1).Visible = True
    lblLabels(2).Visible = True
    lblLabels(2).Caption = "Account Types:"
    txtfields(0).DataField = "Account Name"
    txtfields(1).DataField = "Account ID"
    OpenDB "select [Account Name],[Account ID] from [GL Account Types]"
End Select
  
  If CheckNewDB(ADOprimaryrs, optType(Index).Caption) = True Then
    cmdAdd_Click
  End If
  
  lblTop = optType(Index).Caption
  
End Sub

Public Sub ListType(frm_LIST_Name As String)
  Set db = New ADODB.Connection
  db.CursorLocation = adUseClient
  db.Open gblADOProvider
    
    Select Case frm_LIST_Name
        Case "frm_LIST_Customer_Types"
            optType(0).Value = True
        Case "frm_LIST_Item_Catagories"
            optType(1).Value = True
        Case "frm_LIST_Item_Types"
            optType(6).Value = True
        Case "frm_LIST_Payment_Methods"
            optType(2).Value = True
        Case "frm_LIST_Shipping_Methods"
            optType(3).Value = True
        Case "frm_LIST_Vendor_Types"
            optType(4).Value = True
        Case "frm_LIST_ALL_Types"
            optType(5).Value = True
    End Select
End Sub

Private Sub txtfields_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 1 Then
     keyResponse = CtrlValidate(KeyAscii, "0123456789.")
     If keyResponse = True Then
     Else
        KeyAscii = 0
     End If
End If
End Sub


Private Sub txtFields_LostFocus(Index As Integer)
If Index = 1 Then
    If txtfields(1) = "" Then
        txtfields(1) = "$0.00"
    Else
        txtfields(1) = FormatCurr(txtfields(1))
    End If
End If
End Sub


