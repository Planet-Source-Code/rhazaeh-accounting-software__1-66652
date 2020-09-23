VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_AR_Batch_Posting 
   Caption         =   "Active AR Transaction"
   ClientHeight    =   5205
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   10935
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5205
   ScaleWidth      =   10935
   Begin VB.Frame frPrimary 
      Height          =   4335
      Left            =   0
      TabIndex        =   6
      Top             =   480
      Width           =   10935
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   300
         Left            =   9720
         TabIndex        =   11
         Top             =   2160
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   495
         Left            =   9720
         TabIndex        =   10
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CommandButton cmdPost 
         Caption         =   "&Post"
         Height          =   1020
         Left            =   9720
         Picture         =   "frm_AR_Batch_Posting.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   3120
         Width           =   1095
      End
      Begin MSDataGridLib.DataGrid grdDataGrid 
         Height          =   3975
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   7011
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
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
         Caption         =   "Active AR Transaction"
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "Post YN"
            Caption         =   "Post"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "Yes"
               FalseValue      =   "No"
               NullValue       =   "NA"
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Customer ID"
            Caption         =   "Customer ID"
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
         BeginProperty Column02 
            DataField       =   "Document ID"
            Caption         =   "Document #"
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
         BeginProperty Column03 
            DataField       =   "Document Type"
            Caption         =   "Type"
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
         BeginProperty Column04 
            DataField       =   "Date"
            Caption         =   "Date"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "MM/dd/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "Amount"
            Caption         =   "Amount"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "Status"
            Caption         =   "Status"
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
            MarqueeStyle    =   4
            BeginProperty Column00 
               ColumnWidth     =   705.26
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1094.74
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1500.095
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1094.74
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1305.071
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               ColumnWidth     =   1305.071
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1920.189
            EndProperty
         EndProperty
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
      ScaleWidth      =   10935
      TabIndex        =   0
      Top             =   4905
      Width           =   10935
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   9480
         Picture         =   "frm_AR_Batch_Posting.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   9120
         Picture         =   "frm_AR_Batch_Posting.frx":0784
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Picture         =   "frm_AR_Batch_Posting.frx":0AC6
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Picture         =   "frm_AR_Batch_Posting.frx":0E08
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
         Width           =   8400
      End
   End
   Begin VB.Label lblTop 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Active AR Transaction"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   120
      Width           =   10425
   End
End
Attribute VB_Name = "frm_AR_Batch_Posting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ADOprimaryRS As ADODB.Recordset
Attribute ADOprimaryRS.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean

Dim TempStr As String
Dim db As ADODB.Connection

Private Sub LoadBatchAR()

  'On Error GoTo LoadBatchAR_Error

  'Dim rsBatch As ADODB.Recordset
  'Set rsBatch = New ADODB.Recordset
  'rsBatch.Open "[SYS Purchase Batch]", db, adOpenStatic, adLockOptimistic, adCmdTable
  
  Dim ADOsaleRS As ADODB.Recordset
  Set ADOsaleRS = New ADODB.Recordset
  ADOsaleRS.Open "SELECT [AR SALE Ext Document #],[AR SALE Customer ID],[AR SALE Document Type]," & _
  "[AR SALE Date],[AR SALE Total] FROM [AR Sales] where [AR SALE Posted YN] = false", db, adOpenForwardOnly, adLockOptimistic, adCmdText
  
  db.Execute "DELETE * FROM [SYS Sales Batch]"

  If ADOsaleRS.RecordCount = 0 Then
    MsgBox "There are no unposted purchase transactions."
    Exit Sub
  End If
    
  Dim SQLstatement As String
    
  ADOsaleRS.MoveFirst
  Do While Not ADOsaleRS.EOF
      SQLstatement = "INSERT INTO [SYS Sales Batch]"
      SQLstatement = SQLstatement & " ([Post YN],[Document ID],[Customer ID],[Document Type],[Date],[Amount])"
      SQLstatement = SQLstatement & " VALUES (True,'" & ADOsaleRS("AR SALE Ext Document #") & _
      "','" & ADOsaleRS("AR SALE Customer ID") & "','" & ADOsaleRS("AR SALE Document Type") & _
      "',#" & ADOsaleRS("AR SALE Date") & "#," & ADOsaleRS("AR SALE Total") & ")"
      db.Execute SQLstatement
    
    'rsBatch.AddNew
    '  rsBatch("Post YN") = True
    '  rsBatch("Document ID") = ADOsaleRS("AR SALE Ext Document #")
    '  rsBatch("Vendor ID") = ADOsaleRS("AR SALE Customer ID")
    '  rsBatch("Document Type") = ADOsaleRS("AR SALE Document Type")
    '  rsBatch("Date") = ADOsaleRS("AR SALE Date")
    '  rsBatch("Amount") = ADOsaleRS("AR SALE Total")
    'rsBatch.Update
    ADOsaleRS.MoveNext
  Loop
  
  ADOsaleRS.Close
  Set ADOsaleRS = Nothing
  
  Screen.MousePointer = vbNormal
  Exit Sub
LoadBatchAR_Error:
  Call ErrorLog("AP Batch Posting", "LoadBatchAP", Now, Err.Number, Err.Description, True, db)
  Resume Next

End Sub



Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPost_Click()
Screen.MousePointer = vbHourglass
  'On Error GoTo cmdPost_Click_Error
  
  Dim rsBatch As ADODB.Recordset
  Set rsBatch = New ADODB.Recordset
  rsBatch.Open "SELECT * FROM [SYS Sales Batch] where [Post YN] = True", db, adOpenForwardOnly, adLockOptimistic
'  If rsBatch.BOF = True And rsBatch.EOF = True Then
  If rsBatch.RecordCount = 0 Then
    MsgBox "No transactions are selected to post."
    Screen.MousePointer = vbNormal
    Exit Sub
  End If
  
  Dim X%
  Dim rsSales As ADODB.Recordset
  
  rsBatch.MoveFirst
  Do While Not rsBatch.EOF
    Set rsSales = New ADODB.Recordset
    rsSales.Open "SELECT [AR SALE Document #],[AR SALE Ext Document #],[AR SALE Customer ID],[AR SALE Balance Due]," & _
    "[AR SALE Amount Paid],[AR SALE Check Acct ID],[AR SALE Balance Due],[AR SALE Check Number]," & _
    "[AR SALE Document Type] FROM [AR Sales] WHERE [AR SALE Ext Document #]='" & rsBatch("Document ID") & "'", db, adOpenForwardOnly, adLockOptimistic, adCmdText
    'rsSales.MoveFirst
    'rsSales.Find "[AR SALE Ext Document #]='" & rsBatch("Document ID") & "'"
    
        If rsSales.RecordCount = 0 Then
          MsgBox "Document ID: " & rsBatch("Document ID") & " no longer exists."
          Screen.MousePointer = vbNormal
          Exit Sub
        End If
        X% = Datavalidate(rsSales)
        If X% = False Then
          MsgBox "Error occurred on Document ID: " & rsSales("AR SALE Ext Document #") & Chr$(10) & "Either fix document or do not select to post."
          Screen.MousePointer = vbNormal
          Exit Sub
        End If
  
    rsSales.Close
    Set rsSales = Nothing
    rsBatch.MoveNext
  Loop
  
  Dim Success%
  
  'Now post each entry
  rsBatch.MoveFirst
  Do While Not rsBatch.EOF
    Set rsSales = New ADODB.Recordset
    rsSales.Open "SELECT [AR SALE Document #],[AR SALE Ext Document #],[AR SALE Subtotal]," & _
    "[AR SALE Document Type],[AR SALE Recur Type],[AR SALE Next Recur],[AR SALE Date]," & _
    "[AR SALE Posted YN] FROM [AR Sales] WHERE [AR SALE Ext Document #]='" & rsBatch("Document ID") & _
     "'", db, adOpenForwardOnly, adLockOptimistic, adCmdText
    'rsSales.Seek "=", rsBatch("Document ID")
    'rsSales.MoveFirst
    'rsSales.Find "[AR SALE Ext Document #]='" & rsBatch("Document ID") & "'"
    If rsSales("AR SALE Subtotal") = 0 Then
        MsgBox "Cannot Post a Transaction with $0.00 amount", vbInformation, "Information"
        GoTo JumpLoop
    End If

    db.BeginTrans
    Select Case rsSales("AR SALE Document Type")
    Case "Invoice"
      Success% = PostInvoice(CLng(rsSales("AR SALE Document #")), False, db)
    Case "Return"
      Success% = PostReturn(CLng(rsSales("AR SALE Document #")), False)
    Case "Sales Memo"
      Success% = PostSalesMemo(CLng(rsSales("AR SALE Document #")), False, db)
    Case "Credit Memo"
      Success% = PostCreditMemo(CLng(rsSales("AR SALE Document #")), False)
    End Select
    If Success% = False Then
      db.RollbackTrans
      'rsBatch.Edit
      MsgBox "Transaction Not Posted."
        rsBatch("Status") = "Error"
      rsBatch.Update
    Else
      db.CommitTrans
      'rsBatch.Edit
      MsgBox "Transaction Posted."
        rsBatch("Status") = "Posted"
      rsBatch.Update
      'rsSales.Edit
        Select Case rsSales("AR SALE Recur Type")
        Case "Never"
        Case "Monthly"
          rsSales("AR SALE Next Recur") = DateAdd("m", 1, rsSales("AR SALE Date"))
        Case "Quarterly"
          rsSales("AR SALE Next Recur") = DateAdd("q", 1, rsSales("AR SALE Date"))
        Case "Annually"
          rsSales("AR SALE Next Recur") = DateAdd("yyyy", 1, rsSales("AR SALE Date"))
        End Select
        rsSales("AR SALE Posted YN") = True
      rsSales.Update
    End If
JumpLoop:
    rsSales.Close
    Set rsSales = Nothing
    rsBatch.MoveNext
  Loop
    
  OpenDB
  
    'Display status report
    'DoCmd.OpenReport "rpt - AR Batch Status", acPreview
    rsBatch.Close
    Set rsBatch = Nothing
    
    Screen.MousePointer = vbNormal
  Exit Sub
cmdPost_Click_Error:
  Call ErrorLog("AR Batch Posting", "cmdClick", Now, Err.Number, Err.Description, True, db)
  Resume Next

End Sub

Private Function Datavalidate(rsSales As ADODB.Recordset) As Integer

  'On Error GoTo DataValidate_Error

  'If IsNull(rsSales("AR SALE Customer ID")) Then
  '  MsgBox "Enter a customer!", , "Error"
  '  Datavalidate = False
  '  Exit Function
  'End If

  'If Len(Trim(rsSales("AR SALE Customer ID"))) = 0 Then
  '  MsgBox "Enter a customer!", , "Error"
  '  Datavalidate = False
  '  Exit Function
  'End If
  'Does customer exist
  'Dim rsCustomer As ADODB.Recordset
  'Set rsCustomer = New ADODB.Recordset
  'rsCustomer.Open "SELECT * FROM [AR Customer]", db, adOpenStatic, adLockOptimistic
  'rsCustomer.Index = "PrimaryKey"
  'rsCustomer.Find "[AR CUST Customer ID]='" & rsSales("AR SALE Customer ID") & "'"
  'If rsCustomer.EOF Then
  '  MsgBox "Customer does not exist!", , "Error"
  '  Datavalidate = False
  '  Exit Function
  'End If

  'Check Credit limit for this customer
  Dim Limit#
  Dim CurrentBalance#
  Dim Response%
  Limit# = LookRecord("[AR CUST Credit Limit]", "[AR Customer]", db, "[AR CUST Customer ID] = '" & rsSales("AR SALE Customer ID") & "'")
  If Limit# > 0 Then
    CurrentBalance# = LookRecord("[AR CUST Financial Period 1]", "[AR Customer]", db, "[AR CUST Customer ID] = '" & rsSales("AR SALE Customer ID") & "'")
    CurrentBalance# = CurrentBalance# + rsSales("AR SALE Balance Due")
    If CurrentBalance# > Limit# Then
      Response% = MsgBox("New balance will exceed customers credit limit!" & vbCr & "Credit Limit            : " & FormatCurr(CCur(Limit#)) & vbCr & "Previous Balance : " & FormatCurr(CurrentBalance# - CurrentRequest) & vbCr & "New Request       : " & FormatCurr(CCur(CurrentRequest)) & vbCr & "Current Balance   : " & FormatCurr(CCur(CurrentBalance#)) & vbCr & vbCr & "Would you like to Continue?", vbYesNo, "Information")
      If Response% = vbNo Then
        Datavalidate = False
        Exit Function
      End If
    End If
  End If

  If rsSales("AR SALE Amount Paid") > 0 Then 'Or Me![AR SALE Document Type] = "Credit Memo" Then
    If IsNull(rsSales("AR SALE Check Acct ID")) Then
      MsgBox "Please enter a bank account!", , "Error"
      Datavalidate = False
      Exit Function
    End If

    'Verify bank account
    'Dim tmp$
    'Dim rsGL As ADODB.Recordset
    'Set rsGL = New ADODB.Recordset
    'rsGL.Open "SELECT * FROM [GL Chart Of Accounts]", db, adOpenStatic, adLockOptimistic
    'rsGL.Index = "PrimaryKey"
    'tmp$ = rsSales("AR SALE Check Acct ID")
    'rsGL.Seek "=", tmp$
    'If rsGL.NoMatch Then
    '  MsgBox "Not a valid bank account!", , "Error"
    '  DataValidate = False
    '  Exit Function
    'End If

    'If rsGL("GL COA Asset Type") = "Cash" Then
    'Else
    '  MsgBox "Not a valid back account!", , "Error"
    '  DataValidate = False
    '  Exit Function
    'End If
    
    'Check for balance due < 0 and check number
    If rsSales("AR SALE Balance Due") < 0 Then
      MsgBox "Amount paid cannot exceed invoice total!", , "Error"
      Datavalidate = False
      Exit Function
    End If
    If NZ(rsSales("AR SALE Check Number")) = "" Then
      MsgBox "You must enter a check number!", , "Error"
      Datavalidate = False
      Exit Function
    End If

  End If
  
  If rsSales("AR SALE Document Type") = "Invoice" Then
    If CountRecord("[AR SALED Item ID]", "[AR Sales Detail]", db, "[AR SALED Document #] = " & rsSales("AR SALE Document #")) <= 0 Then
      MsgBox "Must have at least one inventory item!", , "Error"
      Datavalidate = False
      Exit Function
    End If
  End If

  Datavalidate = True
  
  Exit Function
DataValidate_Error:
  Call ErrorLog("AR Batch Posting", "DataValidate", Now, Err.Number, Err.Description, True, db)
  Resume Next
  
End Function

Private Sub cmdRefresh_Click()
    OpenDB
End Sub

Private Sub Form_Load()
Screen.MousePointer = vbHourglass
On Error GoTo FormErr
  Set db = New ADODB.Connection
  db.CursorLocation = adUseClient
  db.Open gblADOProvider
  
  OpenDB
  
  Screen.MousePointer = vbNormal
Exit Sub
FormErr:
  MsgBox Err.Description
  Screen.MousePointer = vbDefault
End Sub

Private Sub OpenDB()
Screen.MousePointer = vbHourglass
  
  LoadBatchAR
  
  Set grdDataGrid.DataSource = Nothing
  
  If ADOprimaryRS Is Nothing Then
  Else
    ADOprimaryRS.Close
  End If
  
  TempStr = "select * from [SYS Sales Batch]"
  
  Set ADOprimaryRS = New ADODB.Recordset
  ADOprimaryRS.Open TempStr, db, adOpenStatic, adLockOptimistic
  ADOprimaryRS.MoveFirst
  Do While ADOprimaryRS.EOF = False
    ADOprimaryRS![Post YN] = False
    If IsNull(ADOprimaryRS![Status]) Or ADOprimaryRS![Status] = "" Then
        ADOprimaryRS![Status] = "Open"
    End If
    ADOprimaryRS.Update
    ADOprimaryRS.MoveNext
  Loop
  
  Set grdDataGrid.DataSource = ADOprimaryRS
  
  grdDataGrid.Columns(0).Button = True
  mbDataChanged = False
  
  GetTextColor Me
End Sub


Private Sub Form_Resize()
  If fMainForm.WindowState = 1 Then Exit Sub
  If Me.WindowState = 0 Then
  ElseIf Me.WindowState = 2 Then
    GoTo SkipResize
  Else
    Exit Sub
  End If
  
  Me.Width = 11055
  Me.Height = 5610
SkipResize:
  frPrimary.Left = (Me.ScaleWidth - frPrimary.Width) / 2
  lblTop.Width = frPrimary.Width
  lblTop.Left = frPrimary.Left
  frPrimary.Top = (Me.ScaleHeight - frPrimary.Height - picStatBox.Height) / 2 + 230
  
  lblStatus.Width = Me.Width - 1500
  cmdNext.Left = lblStatus.Width + 700
  cmdLast.Left = cmdNext.Left + 340
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If mbEditFlag Or mbAddNewFlag Then Exit Sub

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
  Screen.MousePointer = vbHourglass
      If ADOprimaryRS.RecordCount > 0 Then
        If ADOprimaryRS.EditMode <> 0 Then
          ADOprimaryRS.CancelUpdate
        End If
      End If
      ADOprimaryRS.Close
      Set ADOprimaryRS = Nothing
      db.Close
      Set db = Nothing
  Screen.MousePointer = vbDefault
  Set frm_AR_Batch_Posting = Nothing
Exit Sub
FormErr:
  MsgBox Err.Description
  Screen.MousePointer = vbDefault
End Sub

Private Sub adoPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This will display the current record position for this recordset
  lblStatus.Caption = "Record: " & CStr(ADOprimaryRS.AbsolutePosition) & " of " & CStr(ADOprimaryRS.RecordCount)
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

Private Sub cmdFirst_Click()
  'On Error GoTo GoFirstError

  ADOprimaryRS.MoveFirst
  mbDataChanged = False

  Exit Sub

GoFirstError:
  MsgBox Err.Description
End Sub

Private Sub cmdLast_Click()
  'On Error GoTo GoLastError

  ADOprimaryRS.MoveLast
  mbDataChanged = False

  Exit Sub

GoLastError:
  MsgBox Err.Description
End Sub

Private Sub cmdNext_Click()
  'On Error GoTo GoNextError

  If Not ADOprimaryRS.EOF Then ADOprimaryRS.MoveNext
  If ADOprimaryRS.EOF And ADOprimaryRS.RecordCount > 0 Then
    Beep
     'moved off the end so go back
    ADOprimaryRS.MoveLast
  End If
  'show the current record
  mbDataChanged = False

  Exit Sub
GoNextError:
  MsgBox Err.Description
End Sub

Private Sub cmdPrevious_Click()
  'On Error GoTo GoPrevError

  If Not ADOprimaryRS.BOF Then ADOprimaryRS.MovePrevious
  If ADOprimaryRS.BOF And ADOprimaryRS.RecordCount > 0 Then
    Beep
    'moved off the end so go back
    ADOprimaryRS.MoveFirst
  End If
  'show the current record
  mbDataChanged = False

  Exit Sub

GoPrevError:
  MsgBox Err.Description
End Sub

Private Sub grdDataGrid_ButtonClick(ByVal ColIndex As Integer)
    If CCur(grdDataGrid.Columns(5).Value) <= 0 Or grdDataGrid.Columns(6).Value = "Posted" Then
        MsgBox "This transaction cannot be post"
        Exit Sub
    End If
    If grdDataGrid.Columns(0).Text = "No" Then
       grdDataGrid.Columns(0).Text = "Yes"
    Else
       grdDataGrid.Columns(0).Text = "No"
    End If
          SendKeys ("{ENTER}")
          SendKeys ("{down}")
          SendKeys ("{up}")
    'grdDataGrid.Refresh
End Sub

Private Sub grdDataGrid_HeadClick(ByVal ColIndex As Integer)
If ADOprimaryRS.RecordCount = 0 Then Exit Sub
    ADOprimaryRS.Close
    Set ADOprimaryRS = Nothing
    Set ADOprimaryRS = New ADODB.Recordset
    ADOprimaryRS.Open TempStr & " ORDER BY [" & grdDataGrid.Columns(ColIndex).DataField & "]", db, adOpenKeyset, adLockOptimistic, adCmdText
    Set grdDataGrid.DataSource = ADOprimaryRS
End Sub


