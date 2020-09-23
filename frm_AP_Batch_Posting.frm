VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_AP_Batch_Posting 
   Caption         =   "Active AP Transaction"
   ClientHeight    =   7350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14310
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7350
   ScaleWidth      =   14310
   Begin VB.Frame frPrimary 
      Height          =   6855
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   14295
      Begin MSDataGridLib.DataGrid grdDataGrid 
         Height          =   6495
         Left            =   3480
         TabIndex        =   13
         Top             =   240
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   11456
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
         Caption         =   "Active AP Transaction"
         ColumnCount     =   10
         BeginProperty Column00 
            DataField       =   "AP PO Saved YN"
            Caption         =   "Select"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "Yes"
               FalseValue      =   "No"
               NullValue       =   "N/A"
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "AP PO Vendor ID"
            Caption         =   "Vendor ID"
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
            DataField       =   "AP PO Ext Document No"
            Caption         =   "Document No"
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
            DataField       =   "AP PO Document Type"
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
            DataField       =   "AP PO Ordered by"
            Caption         =   "Ordered by"
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
         BeginProperty Column05 
            DataField       =   "AP PO Payment Terms"
            Caption         =   "Terms"
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
         BeginProperty Column06 
            DataField       =   "AP PO Date"
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
         BeginProperty Column07 
            DataField       =   "AP PO Due Date"
            Caption         =   "Due Date"
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
         BeginProperty Column08 
            DataField       =   "AP PO Total Amount"
            Caption         =   "Amount"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """$""#,##0.00;(""$""#,##0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column09 
            DataField       =   "AP PO Status"
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
               ColumnWidth     =   734.74
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1019.906
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1154.835
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   929.764
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   929.764
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1035.213
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1035.213
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1035.213
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1140.095
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame1 
         Height          =   6615
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   3255
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "frm_AP_Batch_Posting.frx":0000
            Left            =   1320
            List            =   "frm_AP_Batch_Posting.frx":0016
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   3240
            Width           =   1815
         End
         Begin VB.CommandButton cmdClose 
            Caption         =   "&Close"
            Height          =   315
            Left            =   2160
            Picture         =   "frm_AP_Batch_Posting.frx":0049
            TabIndex        =   18
            Top             =   6120
            Width           =   975
         End
         Begin VB.Frame Frame3 
            Height          =   1335
            Left            =   120
            TabIndex        =   21
            Top             =   120
            Width           =   3015
            Begin VB.CommandButton cmdSearch 
               Caption         =   "&Search"
               Height          =   855
               Left            =   1920
               Picture         =   "frm_AP_Batch_Posting.frx":0C25
               Style           =   1  'Graphical
               TabIndex        =   23
               Top             =   360
               Width           =   975
            End
            Begin VB.TextBox txtfields 
               BeginProperty DataFormat 
                  Type            =   0
                  Format          =   "MM/dd/yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
               Height          =   285
               Index           =   2
               Left            =   240
               TabIndex        =   22
               Top             =   720
               Width           =   1575
            End
            Begin VB.Label lblLabels 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Document No"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   5
               Left            =   240
               TabIndex        =   24
               Top             =   480
               Width           =   1575
            End
         End
         Begin VB.CommandButton cmdPost 
            Caption         =   "&Post"
            Height          =   795
            Left            =   2160
            Picture         =   "frm_AP_Batch_Posting.frx":0F2F
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   5160
            Width           =   975
         End
         Begin VB.CommandButton cmdShow 
            Caption         =   "S&how Selection"
            Height          =   315
            Left            =   120
            Picture         =   "frm_AP_Batch_Posting.frx":1239
            TabIndex        =   5
            Top             =   2640
            Width           =   3015
         End
         Begin VB.CommandButton cmdALL 
            Caption         =   "&Show All"
            Height          =   315
            Left            =   1200
            TabIndex        =   16
            Top             =   6120
            Width           =   975
         End
         Begin VB.OptionButton optSelect 
            Caption         =   "Unselect All"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   15
            Top             =   4080
            Width           =   1695
         End
         Begin VB.OptionButton optSelect 
            Caption         =   "Select All Ready Document"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   14
            Top             =   3720
            Width           =   2415
         End
         Begin VB.TextBox txtfields 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "MM/dd/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   1560
            Width           =   1455
         End
         Begin VB.CommandButton cmdDate 
            Height          =   285
            Index           =   0
            Left            =   2760
            Picture         =   "frm_AP_Batch_Posting.frx":1543
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   1560
            Width           =   375
         End
         Begin VB.TextBox txtfields 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "MM/dd/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   1920
            Width           =   1455
         End
         Begin VB.CommandButton cmdDate 
            Height          =   285
            Index           =   1
            Left            =   2760
            Picture         =   "frm_AP_Batch_Posting.frx":184D
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   1920
            Width           =   375
         End
         Begin VB.TextBox txtfields 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "MM/dd/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            Height          =   285
            Index           =   35
            Left            =   1320
            TabIndex        =   4
            Top             =   2280
            Width           =   1455
         End
         Begin VB.CommandButton cmdLookupVend 
            Height          =   285
            Left            =   2760
            Picture         =   "frm_AP_Batch_Posting.frx":1B57
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   2280
            Width           =   375
         End
         Begin VB.CommandButton cmdPrint 
            Caption         =   "Prin&t"
            Height          =   795
            Left            =   240
            Picture         =   "frm_AP_Batch_Posting.frx":1E61
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   5160
            Width           =   975
         End
         Begin VB.CommandButton cmdRefresh 
            Caption         =   "&Refresh"
            Height          =   315
            Left            =   240
            TabIndex        =   19
            Top             =   6120
            Width           =   975
         End
         Begin VB.Label lblLabels 
            Caption         =   "Document Type:"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   26
            Top             =   3240
            Width           =   1215
         End
         Begin VB.Label lblLabels 
            Caption         =   "Start Date:"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   12
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Label lblLabels 
            Caption         =   "End Date:"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   11
            Top             =   1920
            Width           =   1215
         End
         Begin VB.Label lblLabels 
            Caption         =   "Bank Account:"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   10
            Top             =   2280
            Width           =   1215
         End
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Active AP Transaction"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   9225
   End
End
Attribute VB_Name = "frm_AP_Batch_Posting"
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

Dim TempStr As String
Dim db As ADODB.Connection
Dim WhichField As String
Dim CriteriaType As String


Private Sub cmdALL_Click()
  TempStr = "SELECT [AP PO Saved YN],[AP PO Vendor ID],[AP PO Ext Document No]," & _
            "[AP PO Document Type],[AP PO Ordered by],[AP PO Payment Terms]," & _
            "[AP PO Date],[AP PO Due Date],[AP PO Total Amount],[AP PO Status], " & _
            "[AP PO Document No],[AP PO Amount Paid],[AP PO Posted YN]  " & _
            "FROM [AP Purchase] WHERE [AP PO Posted YN]=False" & CriteriaType

  OpenDB
End Sub

Private Sub cmdDate_Click(Index As Integer)
Select Case Index
Case 0, 3
    Menu_Calendar.WhoCallMe True, 1302
    'Menu_Calendar.Show vbModal
Case 1, 2
    Menu_Calendar.WhoCallMe True, 1640
    'Menu_Calendar.Show vbModal
End Select
End Sub

Private Sub cmdLookupVend_Click()
    Dim SQLstatement As String
    Dim No As Integer
    Dim ghead As String
    Dim fhead As String

    No = 1220
    SQLstatement = "select [BANK ACCT ID], [BANK ACCT Name]" & _
                    "from [BANK Accounts]"
    ghead = "Account Description"
    fhead = "Account No//Description"
    
    AllLookup.GetWhichTable No, SQLstatement, ghead, fhead, db
End Sub

Private Sub cmdSearch_Click()
If ADOprimaryrs Is Nothing Then
Else
    If ADOprimaryrs.RecordCount = 0 Then Exit Sub
    SearchRECORD ADOprimaryrs, grdDataGrid, txtfields(2).Text, lblLabels(5).Caption, WhichField, "AP PO Ext Document No"
End If

End Sub

Private Sub cmdShow_Click()
    If txtfields(0).Text = "" Or txtfields(1) = "" Then
        MsgBox "You must complete the start and the end date before you could continue", vbInformation, "Information"
        Exit Sub
    End If
  
  If txtfields(35).Text = "" Then
    TempStr = "SELECT [AP PO Saved YN],[AP PO Vendor ID],[AP PO Ext Document No]," & _
              "[AP PO Document Type],[AP PO Ordered by],[AP PO Payment Terms]," & _
              "[AP PO Date],[AP PO Due Date],[AP PO Total Amount],[AP PO Status], " & _
              "[AP PO Document No],[AP PO Amount Paid],[AP PO Posted YN]  " & _
              "FROM [AP Purchase] WHERE [AP PO Posted YN]=False AND [AP PO Date] BETWEEN #" & txtfields(0).Text & "# AND #" & txtfields(1).Text & "#" & CriteriaType
  Else
    TempStr = "SELECT [AP PO Saved YN],[AP PO Vendor ID],[AP PO Ext Document No]," & _
              "[AP PO Document Type],[AP PO Ordered by],[AP PO Payment Terms]," & _
              "[AP PO Date],[AP PO Due Date],[AP PO Total Amount],[AP PO Status], " & _
              "[AP PO Document No],[AP PO Amount Paid],[AP PO Posted YN]  " & _
              "FROM [AP Purchase] WHERE [AP PO Posted YN]=False AND [AP PO Check Acct ID]='" & txtfields(35).Text & "' AND [AP PO Date] BETWEEN #" & txtfields(0).Text & "# AND #" & txtfields(1).Text & "#" & CriteriaType
  End If
  OpenDB

End Sub

Private Sub Combo1_Click()

Select Case Combo1.Text
Case "All"
    CriteriaType = ""
Case "PO"
    CriteriaType = " AND [AP PO Document Type]='" & Combo1.Text & "'"
Case "Receiving"
    CriteriaType = " AND [AP PO Document Type]='" & Combo1.Text & "'"
Case "Credit Memo"
    CriteriaType = " AND [AP PO Document Type]='" & Combo1.Text & "'"
Case "Voucher"
    CriteriaType = " AND [AP PO Document Type]='" & Combo1.Text & "'"
Case "RMA"
    CriteriaType = " AND [AP PO Document Type]='" & Combo1.Text & "'"
Case Else
    CriteriaType = ""
End Select

End Sub

Private Sub Form_Load()
ShowStatus True
On Error GoTo FormErr
  Set db = New ADODB.Connection
  db.CursorLocation = adUseClient
  db.Open gblADOProvider
  
  TempStr = "SELECT [AP PO Saved YN],[AP PO Vendor ID],[AP PO Ext Document No]," & _
            "[AP PO Document Type],[AP PO Ordered by],[AP PO Payment Terms]," & _
            "[AP PO Date],[AP PO Due Date],[AP PO Total Amount],[AP PO Status], " & _
            "[AP PO Document No],[AP PO Amount Paid],[AP PO Posted YN] " & _
            "FROM [AP Purchase] WHERE [AP PO Posted YN]=False" & CriteriaType

  OpenDB
  
  grdDataGrid.Columns(0).Button = True
  mbDataChanged = False
  ShowStatus False
Exit Sub
FormErr:
  MsgBox Err.Description
  ShowStatus False
End Sub

Private Sub OpenDB()
ShowStatus True
  
  Set grdDataGrid.DataSource = Nothing
  
  If ADOprimaryrs Is Nothing Then
  Else
    ADOprimaryrs.Close
  End If
  
  Set ADOprimaryrs = New ADODB.Recordset
  ADOprimaryrs.Open TempStr, db, adOpenStatic, adLockOptimistic
  
  Set grdDataGrid.DataSource = ADOprimaryrs
  
  mbDataChanged = False
  
  GetTextColor Me
End Sub


Private Sub Form_Resize()
  'On Error Resume Next
  'This will resize the grid when the form is resized
  If fMainForm.WindowState = 1 Then Exit Sub
  If Me.WindowState = 0 Then
  ElseIf Me.WindowState = 2 Then
    GoTo SkipResize
  Else
    Exit Sub
  End If
  
  Me.Width = 14430
  Me.Height = 7755
SkipResize:
  frPrimary.Left = (Me.ScaleWidth - frPrimary.Width) / 2
  Label1.Left = frPrimary.Left
  Label1.Width = frPrimary.Width
  frPrimary.Top = (Me.ScaleHeight - frPrimary.Height) / 2 + 230
  'This will resize the grid when the form is resized
  'grdDataGrid.Height = Me.ScaleHeight - 30 - picButtons.Height - picStatBox.Height
  'lblStatus.Width = Me.Width - 1500
  'cmdNext.Left = lblStatus.Width + 700
  'cmdLast.Left = cmdNext.Left + 340
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
  Set frm_AP_Batch_Posting = Nothing
  ShowStatus False
  
Exit Sub
FormErr:
  MsgBox Err.Description
  ShowStatus False
End Sub

Private Sub cmdRefresh_Click()
    OpenDB
End Sub

Private Sub grdDataGrid_ButtonClick(ByVal ColIndex As Integer)
   Select Case ColIndex
   Case 0
      If grdDataGrid.Columns(6).Value = "Ready" Then
        MsgBox "This Document is not ready to be posted"
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
   End Select

End Sub
Private Sub cmdPost_Click()
ShowStatus True

  'On Error GoTo cmdPost_Click_Error
  Set grdDataGrid.DataSource = Nothing
  Dim Success%
  Dim CantPost As Boolean
  
  'Now post each entry
  ADOprimaryrs.MoveFirst
  Do While Not ADOprimaryrs.EOF
    ShowStatus True
    'rsPurchase.Seek "=", rsBatch("Document ID")
    'MsgBox ADOprimaryrs("AP PO Document Type") & " " & ADOprimaryrs![AP PO Amount Paid]
    If ADOprimaryrs![AP PO Amount Paid] > 0 Then
        MsgBox "Can't Post through here [got check to print], use normal procedure"
        CantPost = True
    Else
        CantPost = False
    End If
    
    If CantPost = False And ADOprimaryrs![AP PO Status] = "Ready" Then
        db.BeginTrans
        Select Case ADOprimaryrs("AP PO Document Type")
        Case "PO"
          Success% = ClonePurchase(CLng(ADOprimaryrs("AP PO Document No")), db)
        Case "Receiving"
          Success% = PostReceiving(CLng(ADOprimaryrs("AP PO Document No")), False, db)
        Case "Voucher"
          Success% = PostVoucher(CLng(ADOprimaryrs("AP PO Document No")), False, db)
        Case "Credit Memo"
          Success% = PostPOCreditMemo(CLng(ADOprimaryrs("AP PO Document No")), False, db)
        Case "RMA"
          Success% = PostRMA(CLng(txtfields(2)), True, db)
        Case Else
          Exit Sub
        End Select
        
    
        If Success% = False Then
          db.RollbackTrans
          'rsBatch.Edit
            MsgBox "Transaction Not Posted"
            ADOprimaryrs("AP PO Status") = "Error"
          ADOprimaryrs.Update
        Else
            db.CommitTrans
            'rsPurchase.Edit
            MsgBox "Transaction Posted"
            If ADOprimaryrs("AP PO Document Type") <> "PO" Then
                ADOprimaryrs("AP PO Posted YN") = True
                ADOprimaryrs("AP PO Status") = "Posted"
                ADOprimaryrs.Update
            End If
        End If
        ShowStatus False
    End If
    ADOprimaryrs.MoveNext
  Loop
  
  'ADOprimaryrs.MoveFirst
  cmdRefresh_Click
  Set grdDataGrid.DataSource = ADOprimaryrs
  ShowStatus False
                                                    
  'Display status report
  'DoCmd.OpenReport "rpt - AP Batch Status", acPreview
  Exit Sub
cmdPost_Click_Error:
  Call ErrorLog("AP Batch Posting", "cmdPost", Now, Err.Number, Err.Description, True, db)
  Resume Next
  
End Sub

Private Sub grdDataGrid_HeadClick(ByVal ColIndex As Integer)
If ADOprimaryrs.RecordCount = 0 Then Exit Sub
    lblLabels(5) = grdDataGrid.Columns(ColIndex).Caption
    WhichField = grdDataGrid.Columns(ColIndex).DataField
    ADOprimaryrs.Close
    Set ADOprimaryrs = Nothing
    Set ADOprimaryrs = New ADODB.Recordset
    ADOprimaryrs.Open TempStr & " ORDER BY [" & WhichField & "]", db, adOpenKeyset, adLockOptimistic, adCmdText
    Set grdDataGrid.DataSource = ADOprimaryrs
End Sub

Private Sub Selection(SelectType As Integer)

Set grdDataGrid.DataSource = Nothing
    With ADOprimaryrs
        .MoveFirst
        Do While Not .EOF
            If ![AP PO Status] = "Ready" Then
                ![AP PO Saved YN] = SelectType
            End If
            .MoveNext
        Loop
        .MoveFirst
    End With
Set grdDataGrid.DataSource = ADOprimaryrs

End Sub

Private Sub optSelect_Click(Index As Integer)
    Selection Index
End Sub
