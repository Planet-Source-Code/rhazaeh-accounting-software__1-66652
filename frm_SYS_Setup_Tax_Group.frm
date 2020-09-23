VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_SYS_Setup_Tax_Group 
   Caption         =   "Tax Group"
   ClientHeight    =   4185
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4185
   ScaleWidth      =   7260
   Begin VB.Frame frprimary 
      Height          =   3015
      Left            =   0
      TabIndex        =   21
      Top             =   480
      Width           =   7215
      Begin VB.TextBox txtFields 
         DataField       =   "SYS TAXGRP ID"
         Height          =   285
         Index           =   0
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton cmdLookup 
         Height          =   285
         Left            =   3000
         Picture         =   "frm_SYS_Setup_Tax_Group.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "SYS TAXGRP Name"
         Height          =   285
         Index           =   1
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   720
         Width           =   3375
      End
      Begin VB.Frame Frame2 
         Enabled         =   0   'False
         Height          =   1695
         Left            =   240
         TabIndex        =   22
         Top             =   1200
         Width           =   4455
         Begin VB.TextBox txtTax 
            DataField       =   "SYS TAX ID"
            Height          =   285
            Index           =   0
            Left            =   240
            TabIndex        =   4
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox txtTax 
            DataField       =   "SYS TAX Description"
            Height          =   285
            Index           =   1
            Left            =   1680
            TabIndex        =   5
            Top             =   600
            Width           =   2535
         End
         Begin VB.TextBox txtTax 
            Alignment       =   2  'Center
            DataField       =   "SYS TAX Percent"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   240
            TabIndex        =   6
            Top             =   1320
            Width           =   1215
         End
         Begin VB.TextBox txtTax 
            Alignment       =   2  'Center
            DataField       =   "SYS TAX Account"
            Height          =   285
            Index           =   3
            Left            =   3000
            TabIndex        =   7
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label lblTax 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tax ID"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   26
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label lblTax 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tax Description"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   1680
            TabIndex        =   25
            Top             =   360
            Width           =   2535
         End
         Begin VB.Label lblTax 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tax Percent"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   24
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label lblTax 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tax Account"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   3000
            TabIndex        =   23
            Top             =   1080
            Width           =   1215
         End
      End
      Begin MSDataGridLib.DataGrid grdDataGrid 
         Bindings        =   "frm_SYS_Setup_Tax_Group.frx":014A
         Height          =   2355
         Left            =   4800
         TabIndex        =   27
         Top             =   600
         Width           =   2280
         _ExtentX        =   4022
         _ExtentY        =   4154
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   9164498
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         AllowDelete     =   -1  'True
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
         Caption         =   "Tax"
         ColumnCount     =   1
         BeginProperty Column00 
            DataField       =   "SYS TAXGRPD Tax ID"
            Caption         =   "Tax ID"
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   1739.906
            EndProperty
         EndProperty
      End
      Begin VB.Label lblfields 
         Caption         =   " Group ID:"
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   30
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Name:"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   29
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Click To Select"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   4800
         TabIndex        =   28
         Top             =   360
         Width           =   2280
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
      ScaleWidth      =   7260
      TabIndex        =   14
      Top             =   3885
      Width           =   7260
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Picture         =   "frm_SYS_Setup_Tax_Group.frx":015A
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Picture         =   "frm_SYS_Setup_Tax_Group.frx":049C
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Picture         =   "frm_SYS_Setup_Tax_Group.frx":07DE
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Picture         =   "frm_SYS_Setup_Tax_Group.frx":0B20
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
         TabIndex        =   19
         Top             =   0
         Width           =   3360
      End
   End
   Begin VB.PictureBox PicReport 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   0
      ScaleHeight     =   360
      ScaleWidth      =   7260
      TabIndex        =   0
      Top             =   3525
      Width           =   7260
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Canc&el"
         Height          =   300
         Left            =   4800
         TabIndex        =   8
         Top             =   0
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   300
         Left            =   3840
         TabIndex        =   13
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   300
         Left            =   2880
         TabIndex        =   12
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   300
         Left            =   1920
         TabIndex        =   11
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   300
         Left            =   960
         TabIndex        =   10
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   300
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.Label lblTop 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tax Group"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   20
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "frm_SYS_Setup_Tax_Group"
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
Dim NewLoad As Boolean

Public grdOnAddNew As Boolean

Public Sub CallByUser(GroupID As String)
    Me.Show
    If mbAddNewFlag = False Then
        cmdAdd_Click
        txtFields(0) = GroupID
    Else
        txtFields(0) = GroupID
    End If
End Sub


Private Sub grdDataGrid_BeforeDelete(Cancel As Integer)
If mbAddNewFlag = True Then Exit Sub
    Dim DeleteCration As Integer
    
    DeleteCration = MsgBox("Attempting to delete the data. " & vbCr & "Are you sure?", vbYesNo, "Deleting Confirmation")
    If DeleteCration = vbNo Then Cancel = 1

End Sub

Private Sub cmdLookup_Click()
   AllLookup.ToWhichRecord ADOprimaryrs, "Tax", ""
   'AllLookup.Show vbModal
End Sub

Private Sub Form_Load()
On Error GoTo FormErr
ShowStatus True
  NewLoad = True
  Dim CreateOrder As Integer
  
  Set db = New ADODB.Connection
  db.CursorLocation = adUseClient
  db.Provider = "MSDataShape"
  db.Open "Data " & gblADOProvider

  Set ADOprimaryrs = New ADODB.Recordset
  ADOprimaryrs.Open "SHAPE {select * from [SYS Tax Group]} AS ParentCMD APPEND ({select * from [SYS Tax Group Detail] } AS ChildCMD RELATE [SYS TAXGRP ID] TO [SYS TAXGRPD Group ID]) AS ChildCMD", db, adOpenStatic, adLockOptimistic, adCmdText
  'ADOprimaryrs.Open "select * from [SYS Tax Group] Order by [SYS TAXGRP ID]", db, adOpenStatic, adLockOptimistic, adCmdText
 
  Dim Ctrl As TextBox
  For Each Ctrl In Me.txtFields
        Set Ctrl.DataSource = ADOprimaryrs
        If Ctrl.DataField <> "" Then
            If ADOprimaryrs("" & Ctrl.DataField & "").Type = 202 Then Ctrl.MaxLength = ADOprimaryrs("" & Ctrl.DataField & "").DefinedSize
        End If
  Next
            
  Me.Width = 7350
  Me.Height = 4560
    
    grdOnAddNew = False
    grdDataGrid.Columns(0).Button = True
    grdDataGrid.AllowAddNew = True
    grdDataGrid.AllowDelete = True
  
  If CheckNewDB(ADOprimaryrs, "Tax Group") = True Then
    cmdAdd_Click
  Else
    Set grdDataGrid.DataSource = ADOprimaryrs("ChildCMD").UnderlyingValue
  End If
    
    GetTextColor Me
    
    mbDataChanged = False
    NewLoad = False
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
  
  Me.Width = 7350
  Me.Height = 4560
  
SkipResize:
  frPrimary.Left = (Me.ScaleWidth - frPrimary.Width) / 2
  lblTop.Left = frPrimary.Left
  lblTop.Width = frPrimary.Width
  frPrimary.Top = (Me.ScaleHeight - frPrimary.Height - PicReport.Height - picStatBox.Height) / 2 + 230
  
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
      EndLoad db, ADOprimaryrs, "Tax Groups"
      If ADOprimaryrs.RecordCount > 0 Then
        If ADOprimaryrs.EditMode <> 0 Then
          ADOprimaryrs.CancelUpdate
        End If
      End If
      ADOprimaryrs.Close
      Set ADOprimaryrs = Nothing
      db.Close
      Set db = Nothing
  Set frm_SYS_Setup_Tax_Group = Nothing
  ShowStatus False
Exit Sub
FormErr:
  MsgBox Err.Description
  ShowStatus False
End Sub

Private Sub adoPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  If ADOprimaryrs.BOF Or ADOprimaryrs.EOF Then Exit Sub
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
'  'On Error GoTo AddErr
ShowStatus True
If cmdAdd.Caption = "&Save" Then
     If Not CheckEmpty Then
        ShowStatus False
        Exit Sub
     End If
     With ADOprimaryrs
         .Update
         '.MovePrevious
         'grdDataGrid.HoldFields
         'grdDataGrid.ReBind
         'grdDataGrid.Refresh
         'cmdRefresh_Click
         cmdRefresh_Click
         .MoveLast
         Set grdDataGrid.DataSource = ADOprimaryrs("ChildCMD").UnderlyingValue
         NewLoad = False
     End With
     cmdAdd.Caption = "&Add"
     cmdLookup.Enabled = True
     mbAddNewFlag = False
     SetButtons True

Else
    mbAddNewFlag = True
  With ADOprimaryrs
    If Not (.BOF Or .EOF) Then
      mvBookMark = .Bookmark
    End If
     NewLoad = True
     cmdLookup.Enabled = False
     
     Dim oText As TextBox
     For Each oText In Me.txtTax
        Set oText.DataSource = Nothing
        oText.Text = ""
     Next
    Set grdDataGrid.DataSource = Nothing
         grdDataGrid.HoldFields
         grdDataGrid.ReBind
         grdDataGrid.Refresh
    .AddNew
    lblStatus.Caption = "Add record"
    SetButtons False
  End With
End If
  GetTextColor Me
  ShowStatus False
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
  'On Error GoTo DeleteErr
     If Not DataDelete(ADOprimaryrs, Me, True) Then
     End If
End Sub

Private Sub cmdRefresh_Click()
  'This is only needed for multi user apps
  'On Error GoTo RefreshErr
  Set grdDataGrid.DataSource = Nothing
  'MsgBox adoPrimaryRS.State
  If ADOprimaryrs.State <> 0 Then
      ADOprimaryrs.UpdateBatch adAffectAll
      ADOprimaryrs.Requery
  Else
      ADOprimaryrs.Requery
  End If
  Set grdDataGrid.DataSource = ADOprimaryrs("ChildCMD").UnderlyingValue
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub

Private Sub cmdCancel_Click()
  ShowStatus True
  SetButtons True
  cmdAdd.Caption = "&Add"
  cmdCancel.Visible = False
  cmdLookup.Enabled = True
  mbEditFlag = False
  mbAddNewFlag = False
  ADOprimaryrs.CancelUpdate
  NewLoad = False
  Set grdDataGrid.DataSource = ADOprimaryrs("ChildCMD").UnderlyingValue
  If ADOprimaryrs.RecordCount > 0 Then
    ADOprimaryrs.MoveLast
    ADOprimaryrs.Resync adAffectCurrent
    'Set grdDataGrid.DataSource = adoPrimaryRS("ChildCMD").UnderlyingValue
    'While adoPrimaryRS.State  = adStateExecuting
    'Wend
    'cmdRefresh_Click
    If mvBookMark > 0 Then
      ADOprimaryrs.Bookmark = mvBookMark
    Else
      ADOprimaryrs.MoveFirst
    End If
  End If
  mbDataChanged = False
  GetTextColor Me
  ShowStatus False
End Sub

Private Sub cmdUpdate_Click()
  'On Error GoTo UpdateErr

  ADOprimaryrs.UpdateBatch adAffectAll

  If mbAddNewFlag Then
    ADOprimaryrs.MoveLast              'move to the new record
  End If

  mbEditFlag = False
  mbAddNewFlag = False
  SetButtons True
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
        
    txtFields(0).Locked = bVal
    txtFields(1).Locked = bVal
End Sub


Private Sub grdDataGrid_Error(ByVal DataError As Integer, Response As Integer)
    If DataGridKnownError(DataError) Then
        Response = 0
    End If
End Sub

Private Sub grdDataGrid_GotFocus()
    If mbAddNewFlag = True And Trim(txtFields(0).Text) <> "" Then
'        cmdAdd.SetFocus
        CreateOrder = MsgBox("This Request will save the data to the database? Are sure to continue", vbYesNo, "Save Quote")
        If CreateOrder = vbNo Then Exit Sub
        cmdAdd_Click
    End If
End Sub

Private Sub grdDataGrid_OnAddNew()
    grdOnAddNew = True
    'grdDataGrid.Row = 1
End Sub

Private Sub grdDataGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If NewLoad = True Or grdOnAddNew = True Then Exit Sub
    If ADOprimaryrs.BOF Or ADOprimaryrs.EOF Then Exit Sub
    'MsgBox grdDataGrid.Row
    If grdOnAddNew = False And grdDataGrid.Row > -1 Then
        Dim oText As TextBox
        Dim ADOtaxrs As ADODB.Recordset
        
        Set ADOtaxrs = New ADODB.Recordset
        ADOtaxrs.Open "SELECT * FROM [SYS Tax] WHERE [SYS TAX ID]='" & grdDataGrid.Columns(0).Text & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
        For Each oText In Me.txtTax
            Set oText.DataSource = ADOtaxrs
        Next
        ADOtaxrs.Close
        Set ADOtaxrs = Nothing
    End If
    grdOnAddNew = False
Exit Sub
Damn_Attempt:
     grdDataGrid.AllowUpdate = False
     grdDataGrid.col = 0
exit_sub:
End Sub

Private Sub grdDataGrid_AfterColEdit(ByVal ColIndex As Integer)
If mbAddNewFlag = True Then Exit Sub
    If grdDataGrid.Row = -1 Or grdDataGrid.Columns(0) = "" Then Exit Sub
      SendKeys ("{ENTER}")
  If grdDataGrid.Row > 0 Then
      SendKeys ("{up}")
      SendKeys ("{down}")
  ElseIf grdDataGrid.Row = 0 Then
      SendKeys ("{down}")
      SendKeys ("{up}")
  End If
  ADOprimaryrs.Update
End Sub

Private Sub NewgrdDatagrid()
    NewLoad = True
    NewRowForDataGrid ADOprimaryrs, grdDataGrid, "SYS TAXGRP ID", txtFields(0).Text
    grdOnAddNew = False
    NewLoad = False
End Sub

Private Sub grdDataGrid_ButtonClick(ByVal ColIndex As Integer)
On Error GoTo Error_ButtClick
If mbAddNewFlag = True Then Exit Sub
If grdDataGrid.Columns(0) <> "" Then grdOnAddNew = False
Select Case ColIndex
Case 0   'select the item from the ITEM_ID
    TAX_ITEM
End Select

If grdOnAddNew = True And grdDataGrid.Columns(0) <> "" Then
    NewgrdDatagrid
Else
    grdOnAddNew = False
End If
grdDataGrid_AfterColEdit 0
Exit Sub
Error_ButtClick:
    MsgBox "Please click the Table box before clicking the button"
End Sub

Private Sub TAX_ITEM()
   AllLookup.GetWhichTable 1660, "SELECT [SYS TAX ID], [SYS TAX Description]," & _
   "[SYS TAX Percent] FROM [SYS Tax]", "Product", _
   "Tax ID//Tax Description//Percentage", db
   'AllLookup.Show vbModal

End Sub

Private Function CheckEmpty() As Boolean
    If Trim(txtFields(0).Text) = "" Then
        MsgBox lblLabels(0).Caption & " must be filled", vbCritical, "Error"
        CheckEmpty = False
        Exit Function
    ElseIf Trim(txtFields(1).Text) = "" Then
        MsgBox lblLabels(1).Caption & " must be filled", vbCritical, "Error"
        CheckEmpty = False
        Exit Function
    End If
    CheckEmpty = True
End Function
