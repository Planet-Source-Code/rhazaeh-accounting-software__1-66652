VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_SYS_Setup_Chart_Of_Accounts 
   Caption         =   "Setup Chart Of Accounts"
   ClientHeight    =   5310
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   13410
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5310
   ScaleWidth      =   13410
   Begin VB.Frame frPrimary 
      Height          =   4815
      Left            =   0
      TabIndex        =   10
      Top             =   480
      Width           =   13335
      Begin VB.Frame Frame1 
         Height          =   4575
         Left            =   8520
         TabIndex        =   11
         Top             =   120
         Width           =   4695
         Begin VB.PictureBox picStatBox 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   120
            ScaleHeight     =   300
            ScaleWidth      =   3330
            TabIndex        =   41
            Top             =   4200
            Width           =   3330
            Begin VB.CommandButton cmdFirst 
               Height          =   300
               Left            =   0
               Picture         =   "frm_SYS_Setup_Chart_Of_Accounts.frx":0000
               Style           =   1  'Graphical
               TabIndex        =   45
               Top             =   0
               UseMaskColor    =   -1  'True
               Width           =   345
            End
            Begin VB.CommandButton cmdPrevious 
               Height          =   300
               Left            =   345
               Picture         =   "frm_SYS_Setup_Chart_Of_Accounts.frx":0342
               Style           =   1  'Graphical
               TabIndex        =   44
               Top             =   0
               UseMaskColor    =   -1  'True
               Width           =   345
            End
            Begin VB.CommandButton cmdNext 
               Height          =   300
               Left            =   2555
               Picture         =   "frm_SYS_Setup_Chart_Of_Accounts.frx":0684
               Style           =   1  'Graphical
               TabIndex        =   43
               Top             =   0
               UseMaskColor    =   -1  'True
               Width           =   345
            End
            Begin VB.CommandButton cmdLast 
               Height          =   300
               Left            =   2880
               Picture         =   "frm_SYS_Setup_Chart_Of_Accounts.frx":09C6
               Style           =   1  'Graphical
               TabIndex        =   42
               Top             =   0
               UseMaskColor    =   -1  'True
               Width           =   345
            End
            Begin VB.Label lblStatus 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   720
               TabIndex        =   46
               Top             =   0
               Width           =   1800
            End
         End
         Begin VB.PictureBox picButtons 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   660
            Left            =   120
            ScaleHeight     =   660
            ScaleWidth      =   3330
            TabIndex        =   34
            Top             =   3480
            Width           =   3330
            Begin VB.CommandButton cmdCancel 
               Caption         =   "Canc&el"
               Height          =   300
               Left            =   2160
               TabIndex        =   40
               Top             =   360
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.CommandButton cmdClose 
               Caption         =   "&Close"
               Height          =   300
               Left            =   1080
               TabIndex        =   35
               Top             =   360
               Width           =   1095
            End
            Begin VB.CommandButton cmdDelete 
               Caption         =   "&Delete"
               Height          =   300
               Left            =   2160
               TabIndex        =   37
               Top             =   0
               Width           =   1095
            End
            Begin VB.CommandButton cmdUpdate 
               Caption         =   "&Update"
               Height          =   300
               Left            =   1080
               TabIndex        =   38
               Top             =   0
               Width           =   1095
            End
            Begin VB.CommandButton cmdAdd 
               Caption         =   "&Add"
               Height          =   300
               Left            =   0
               TabIndex        =   39
               Top             =   0
               Width           =   1095
            End
            Begin VB.CommandButton cmdRefresh 
               Caption         =   "&Refresh"
               Height          =   300
               Left            =   0
               TabIndex        =   36
               Top             =   360
               Width           =   1095
            End
         End
         Begin VB.Frame Frame4 
            Height          =   1095
            Left            =   1920
            TabIndex        =   32
            Top             =   120
            Width           =   2655
            Begin VB.TextBox txtfields 
               Alignment       =   2  'Center
               DataField       =   "AR SALE Ship Date"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "MM/dd/yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   3
               EndProperty
               Height          =   285
               Index           =   2
               Left            =   120
               TabIndex        =   0
               Top             =   600
               Width           =   1575
            End
            Begin VB.CommandButton cmdSearch 
               Caption         =   "&Search"
               Height          =   705
               Left            =   1800
               Picture         =   "frm_SYS_Setup_Chart_Of_Accounts.frx":0D08
               Style           =   1  'Graphical
               TabIndex        =   1
               Top             =   240
               Width           =   735
            End
            Begin VB.Label lblLabels 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Account No"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   5
               Left            =   120
               TabIndex        =   33
               Top             =   360
               Width           =   1575
            End
         End
         Begin VB.TextBox txtCOA 
            DataField       =   "GL COA Account Balance"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$""#,##0.00;(""$""#,##0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   2
            EndProperty
            DataSource      =   "adoprimaryrs"
            Height          =   285
            Index           =   2
            Left            =   1920
            TabIndex        =   6
            Top             =   3000
            Width           =   1455
         End
         Begin VB.TextBox txtCOA 
            DataField       =   "GL COA Account Name"
            DataSource      =   "adoprimaryrs"
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   3
            Top             =   1560
            Width           =   3255
         End
         Begin VB.TextBox txtCOA 
            DataField       =   "GL COA Account No"
            DataSource      =   "adoprimaryrs"
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   2
            Top             =   720
            Width           =   1695
         End
         Begin VB.CommandButton BalBut 
            Caption         =   "Balances"
            Height          =   825
            Left            =   3600
            Picture         =   "frm_SYS_Setup_Chart_Of_Accounts.frx":1012
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   1320
            Width           =   975
         End
         Begin VB.CommandButton HistBut 
            Caption         =   "History"
            Height          =   825
            Left            =   3600
            Picture         =   "frm_SYS_Setup_Chart_Of_Accounts.frx":131C
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   2160
            Width           =   975
         End
         Begin VB.CommandButton cmdCOA 
            Caption         =   "Load  COA"
            Height          =   825
            Left            =   3600
            Picture         =   "frm_SYS_Setup_Chart_Of_Accounts.frx":1626
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   3000
            Width           =   975
         End
         Begin VB.ComboBox cbCOA 
            DataField       =   "GL COA Balance Type"
            Height          =   315
            Index           =   1
            ItemData        =   "frm_SYS_Setup_Chart_Of_Accounts.frx":1930
            Left            =   120
            List            =   "frm_SYS_Setup_Chart_Of_Accounts.frx":193D
            TabIndex        =   4
            Text            =   "cbCOA"
            Top             =   2280
            Width           =   1695
         End
         Begin VB.TextBox txtCOA 
            DataField       =   "GL COA Asset Type"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$""#,##0.00;(""$""#,##0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   2
            EndProperty
            DataSource      =   "adoprimaryrs"
            Height          =   285
            Index           =   3
            Left            =   120
            TabIndex        =   5
            Top             =   3000
            Width           =   1455
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Balance"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   1920
            TabIndex        =   16
            Top             =   2760
            Width           =   1455
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Balance Type"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   15
            Top             =   2040
            Width           =   1695
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Account Type"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   14
            Top             =   2760
            Width           =   1455
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Account Name"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   13
            Top             =   1320
            Width           =   3255
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Account No"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   12
            Top             =   480
            Width           =   1695
         End
      End
      Begin MSDataGridLib.DataGrid grdDataGrid 
         Bindings        =   "frm_SYS_Setup_Chart_Of_Accounts.frx":1955
         Height          =   4455
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   8385
         _ExtentX        =   14790
         _ExtentY        =   7858
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         AllowAddNew     =   -1  'True
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
         Caption         =   "Chart Of Accounts"
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "GL COA Account No"
            Caption         =   "Account No"
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
         BeginProperty Column01 
            DataField       =   "GL COA Account Name"
            Caption         =   "Account Name"
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
            DataField       =   "GL COA Asset Type"
            Caption         =   "Account Type"
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
            DataField       =   "GL COA Account Balance"
            Caption         =   "Balance"
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
         BeginProperty Column04 
            DataField       =   "GL COA Balance Type"
            Caption         =   "Balance Type"
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
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2204.788
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1995.024
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1395.213
            EndProperty
         EndProperty
      End
      Begin VB.Frame frCOA 
         Height          =   4575
         Left            =   120
         TabIndex        =   18
         Top             =   120
         Visible         =   0   'False
         Width           =   8415
         Begin VB.Frame Frame3 
            Caption         =   "Business Type"
            Height          =   2175
            Left            =   120
            TabIndex        =   25
            Top             =   1680
            Width           =   3135
            Begin VB.OptionButton optBussType 
               Caption         =   "C Corp"
               Height          =   255
               Index           =   0
               Left            =   600
               TabIndex        =   28
               Top             =   600
               Value           =   -1  'True
               Width           =   1935
            End
            Begin VB.OptionButton optBussType 
               Caption         =   "S Corp"
               Height          =   255
               Index           =   1
               Left            =   600
               TabIndex        =   27
               Top             =   1080
               Width           =   1935
            End
            Begin VB.OptionButton optBussType 
               Caption         =   "Partnership"
               Height          =   255
               Index           =   2
               Left            =   600
               TabIndex        =   26
               Top             =   1560
               Width           =   1935
            End
         End
         Begin VB.ListBox lstBussType 
            Height          =   1815
            Left            =   3600
            TabIndex        =   24
            Top             =   2040
            Width           =   3015
         End
         Begin VB.CommandButton cmdLoadCOA 
            Caption         =   "Load COA"
            Height          =   975
            Left            =   7080
            Picture         =   "frm_SYS_Setup_Chart_Of_Accounts.frx":1970
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   1800
            Width           =   1095
         End
         Begin VB.CommandButton cmdCancelCOA 
            Caption         =   "Back"
            Height          =   975
            Left            =   7080
            Picture         =   "frm_SYS_Setup_Chart_Of_Accounts.frx":1C7A
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   2880
            Width           =   1095
         End
         Begin VB.Frame Frame2 
            Height          =   1335
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   8175
            Begin VB.CommandButton cmdDefault 
               Caption         =   "Load Default"
               Height          =   975
               Left            =   6960
               Picture         =   "frm_SYS_Setup_Chart_Of_Accounts.frx":1F84
               Style           =   1  'Graphical
               TabIndex        =   20
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label2 
               Caption         =   $"frm_SYS_Setup_Chart_Of_Accounts.frx":228E
               Height          =   855
               Left            =   360
               TabIndex        =   21
               Top             =   360
               Width           =   6255
            End
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000001&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Select industry to load Chart Of Accounts"
            ForeColor       =   &H0080FFFF&
            Height          =   255
            Index           =   0
            Left            =   3600
            TabIndex        =   30
            Top             =   1800
            Width           =   3015
         End
         Begin VB.Label lblBussType 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            TabIndex        =   29
            Top             =   360
            Width           =   8175
         End
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Chart Of Accounts (COA)"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   31
      Top             =   120
      Width           =   12945
   End
End
Attribute VB_Name = "frm_SYS_Setup_Chart_Of_Accounts"
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

Dim WhichField As String

Public Sub CallByUserCOA(acctNO As String, Optional Accttype As String)
    Me.Show
    If mbAddNewFlag = False Then
        cmdAdd_Click
        txtCOA(0) = acctNO
        txtCOA(1) = Accttype
    Else
        txtCOA(0) = acctNO
        txtCOA(1) = Accttype
    End If
End Sub

Private Sub adoPrimaryRS_RecordChangeComplete(ByVal adReason As ADODB.EventReasonEnum, _
                                              ByVal cRecords As Long, _
                                              ByVal pError As ADODB.Error, _
                                              adStatus As ADODB.EventStatusEnum, _
                                              ByVal pRecordset As ADODB.Recordset)
'Code that refreshes controls in form

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
    grdDataGrid.ReBind
    grdDataGrid.Refresh
  End Select
End Sub

Private Sub adoPrimaryRS_WillMove(ByVal adReason As ADODB.EventReasonEnum, _
                                  adStatus As ADODB.EventStatusEnum, _
                                  ByVal pRecordset As ADODB.Recordset)

Select Case adReason
Case adRsnAddNew
Case adRsnDelete
Case adRsnUpdate
Case adRsnUndoUpdate
Case adRsnUndoAddNew
Case adRsnUndoDelete
Case adRsnRequery
Case adRsnResynch
Case adRsnClose
Case adRsnMove
Case adRsnFirstChange
Case adRsnMoveFirst
Case adRsnMoveNext
Case adRsnMovePrevious
Case adRsnMoveLast
End Select

End Sub

Private Sub BalBut_Click()
    If IsNull(grdDataGrid.Columns(0).Value) Then
        MsgBox "Invalid Account Number!", vbCritical, "Invalid Account Number"
        Exit Sub
    Else
        frm_GL_Account_Balances.OpenAccount grdDataGrid.Columns(0).Text
        frm_GL_Account_Balances.Show
    End If
End Sub


Private Sub cbCOA_KeyPress(Index As Integer, KeyAscii As Integer)
    keyResponse = CtrlValidate(KeyAscii, "")
    If keyResponse = True Then
    Else
       KeyAscii = 0
    End If
End Sub

Private Sub cbCOA_LostFocus(Index As Integer)
   If CheckCombo(cbCOA(Index)) Then
        MsgBox "There is no such selection", vbInformation, "Information"
   End If
End Sub

Private Sub cmdCancel_Click()

    'ADOprimaryrs.CancelUpdate
    GetTextColor Me
    SetButtons True
    mbAddNewFlag = False
End Sub

Private Sub cmdCancelCOA_Click()
    frCOA.Visible = False
    SetButtons True, 1
End Sub

Private Sub cmdCOA_Click()
  
  'Make sure no accounts have a balance first
  Dim rs As ADODB.Recordset
  Dim X%
  
  ShowStatus True
  
  For X% = 1 To 13
    Set rs = New ADODB.Recordset
    rs.Open "SELECT [GL COA Account No] FROM [GL Chart Of Accounts] where [GL COA CY Period " & Trim(CStr(X%)) & " Amt] > 0", db, adOpenKeyset, adLockOptimistic, adCmdText
        If rs.RecordCount > 0 Then
          MsgBox "One or more accounts have current balances and can not be deleted.", vbCritical, "Load Chart Of Accounts"
          ShowStatus False
          Exit Sub
        End If
  rs.Close
  Set rs = Nothing
  Next X%
  
  Set rs = New ADODB.Recordset
  rs.Open "SELECT [GL COA Account No] FROM [GL Chart Of Accounts] where [GL COA CY Beginning Amt] > 0", db, adOpenKeyset, adLockOptimistic, adCmdText
  If rs.RecordCount > 0 Then
      MsgBox "One or more accounts have current balances and can not be deleted.", vbCritical, "Load Chart Of Accounts"
      ShowStatus False
      Exit Sub
  End If
    
    frCOA.Visible = True
    frCOA.ZOrder 0
    
    optBussType(1).Value = True
    SetButtons False, 3
    
    ShowStatus False
End Sub

Private Sub cmdDefault_Click()
Dim TextCOA As TextBox
    
    ADOprimaryrs.Close
    Set ADOprimaryrs = Nothing

    For Each TextCOA In Me.txtCOA
      Set TextCOA.DataSource = Nothing
    Next
    Set cbCOA(1).DataSource = Nothing
    
    Set grdDataGrid.DataSource = Nothing
    grdDataGrid.Refresh
    grdDataGrid.ReBind
    
    db.Execute "DROP TABLE [GL Chart Of Accounts]"
    db.Execute "SELECT * INTO [GL Chart Of Accounts] FROM [COA]"
    OpenDB
    
    frCOA.Visible = False
    SetButtons True, 1

End Sub

Private Sub cmdLoadCOA_Click()

    If lstBussType.Text = "" Then
        MsgBox "Please select an industry first.", , "Load Chart Of Accounts"
        Exit Sub
    End If

    Set grdDataGrid.DataSource = Nothing
    Dim oText As TextBox
    For Each oText In Me.txtCOA
      Set oText.DataSource = Nothing
    Next

  'Load this industries COA's
  
  Dim X%
  Dim Period$

    Dim Response%
    Response% = MsgBox("This will delete your current chart of accounts.  Continue?", vbYesNo, "Load Chart Of Accounts")
    If Response% = vbNo Then Exit Sub

  ShowStatus True

  db.Execute "DELETE * FROM [GL Chart Of Accounts]"

  Dim rsGLCOA As ADODB.Recordset
  Set rsGLCOA = New ADODB.Recordset
  rsGLCOA.Open "[GL Chart Of Accounts]", db, adOpenKeyset, adLockOptimistic, adCmdTable

  Dim Business As String
  
  Business = lblBussType.Caption
    
  Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
  rs.Open "SELECT * FROM [GL Industry Chart Of Accounts] where [GL COA Industry] = '" & Trim(lstBussType.Text) & " - " & Business$ & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
  'aaa = rs.RecordCount
    Do While Not rs.EOF
      rsGLCOA.AddNew
        rsGLCOA("GL COA Account No") = rs("GL COA Account No") & ""
        rsGLCOA("GL COA Account Name") = rs("GL COA Account Name") & ""
        rsGLCOA("GL COA Asset Type") = rs("GL COA Asset Type") & ""
        rsGLCOA("GL COA Balance Type") = rs("GL COA Balance Type") & ""
        'rsGLCOA("GL COA Notes") = rs("GL COA Notes") & ""
        rsGLCOA("GL COA Report Level") = rs("GL COA Report Level")
        rsGLCOA("GL COA Reporting Level") = rs("GL COA Reporting Level")
        rsGLCOA("GL COA Account Balance") = 0
        rsGLCOA("GL COA CY Beginning Amt") = 0
        For X% = 1 To 14
          Period$ = CStr(X%)
          rsGLCOA("GL COA CY Period " & Period$ & " Amt") = 0
        Next X%
        rsGLCOA("GL COA BUD Beginning Amt") = 0
        For X% = 1 To 13
          Period$ = CStr(X%)
          rsGLCOA("GL COA BUD Period " & Period$ & " Amt") = 0
        Next X%
        rsGLCOA("GL COA PY Beginning Amt") = 0
        For X% = 1 To 13
          Period$ = CStr(X%)
          rsGLCOA("GL COA PY Period " & Period$ & " Amt") = 0
        Next X%
        rsGLCOA("GL COA Inactive YN") = False
      rsGLCOA.Update
      rs.MoveNext
    Loop
    
  rsGLCOA.Close
  rs.Close
  Set rsGLCOA = Nothing
  Set rs = Nothing
  
  Call CreateBankCards

  ShowStatus False

    
    ADOprimaryrs.Close
    Set ADOprimaryrs = Nothing
    
    OpenDB
    
    frCOA.Visible = False
    SetButtons True, 1
Exit Sub

Load_Error:
  MsgBox Error, , "Load COA"
  Resume Next
  
cmdLoadGL_Click_Error:
  Call ErrorLog("Load Chart Of Accounts", "cmdLoadGL_Click", Now, Err.Number, Err.Description, True, db)
  Resume Next

End Sub

Private Sub cmdSearch_Click()
    SearchRECORD ADOprimaryrs, grdDataGrid, txtFields(2).Text, lblLabels(5).Caption, WhichField, "GL COA Account No"
End Sub

Private Sub Form_Load()
ShowStatus True
On Error GoTo FormErr
    
    Set db = New ADODB.Connection
    db.CursorLocation = adUseClient
    db.Open gblADOProvider

  OpenDB
  grdDataGrid.AllowAddNew = False
  mbDataChanged = False
  
  GetTextColor Me
ShowStatus False
Exit Sub
FormErr:
  MsgBox Err.Description
  ShowStatus False
End Sub

Private Sub OpenDB()
  Set ADOprimaryrs = New ADODB.Recordset
  ADOprimaryrs.Open "select [GL COA Account No],[GL COA Account Name]," & _
    "[GL COA Asset Type],[GL COA Account Balance],[GL COA Balance Type] " & _
    "from [GL Chart Of Accounts] Order by [GL COA Account No]", db, adOpenStatic, adLockOptimistic
  If ADOprimaryrs.RecordCount = 0 Then
    cmdAdd.Enabled = True
    cmdUpdate.Enabled = False
    cmdDelete.Enabled = False
    cmdRefresh.Enabled = False
    BalBut.Enabled = False
    HistBut.Enabled = False
  End If
SetDataSource
End Sub

Private Sub SetDataSource()
  Dim oText As TextBox
  For Each oText In Me.txtCOA
    Set oText.DataSource = ADOprimaryrs
    If ADOprimaryrs("" & oText.DataField & "").Type = 202 Then oText.MaxLength = ADOprimaryrs("" & oText.DataField & "").DefinedSize
  Next
  
  Dim cbText As ComboBox
  For Each cbText In Me.cbCOA
    Set cbText.DataSource = ADOprimaryrs
  Next
  
  Set grdDataGrid.DataSource = ADOprimaryrs
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
  
  Me.Width = 13530
  Me.Height = 5715
  
SkipResize:
  frPrimary.Left = (Me.ScaleWidth - frPrimary.Width) / 2
  Label1(2).Left = frPrimary.Left
  Label1(2).Width = frPrimary.Width
  frPrimary.Top = (Me.ScaleHeight - frPrimary.Height) / 2 + 230
  
  'lblStatus.Width = Me.Width - 1500
  'cmdNext.Left = lblStatus.Width + 700
  'cmdLast.Left = cmdNext.Left + 340
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
If mbAddNewFlag = True Then cmdCancel_Click
On Error GoTo FormErr
  ShowStatus True
      EndLoad db, ADOprimaryrs, "Chart Of Accounts"
      If ADOprimaryrs.RecordCount > 0 Then
        If ADOprimaryrs.EditMode <> 0 Then
          ADOprimaryrs.CancelUpdate
        End If
      End If
      ADOprimaryrs.Close
      Set ADOprimaryrs = Nothing
      db.Close
      Set db = Nothing
  Set frm_SYS_Setup_Chart_Of_Accounts = Nothing
  ShowStatus False
Exit Sub
FormErr:
  MsgBox Err.Description
  ShowStatus False
End Sub

Private Sub adoPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, _
                                      ByVal pError As ADODB.Error, _
                                      adStatus As ADODB.EventStatusEnum, _
                                      ByVal pRecordset As ADODB.Recordset)
'On Error Resume Next
  'This will display the current record position for this recordset
  If Not adStatus = adStatusErrorsOccurred Then
    lblStatus.Caption = "Record: " & CStr(ADOprimaryrs.AbsolutePosition) & _
        " of " & ADOprimaryrs.RecordCount
  End If

End Sub

Private Sub adoPrimaryRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, _
                                          ByVal cRecords As Long, _
                                          adStatus As ADODB.EventStatusEnum, _
                                          ByVal pRecordset As ADODB.Recordset)
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
    If Not grdDataGrid.Row = -1 Then
        If grdDataGrid.Columns(3).Text = "" Then
            grdDataGrid.Columns(3).Text = "0"
        End If
    End If
  End Select

  If bCancel Then adStatus = adStatusCancel
End Sub

Private Sub cmdAdd_Click()
  On Error GoTo AddErr
    mbAddNewFlag = True
  With ADOprimaryrs
    .AddNew
    lblStatus.Caption = "Add record"
    txtCOA(0).Enabled = True
    txtCOA(1).Enabled = True
    txtCOA(3).Enabled = True
    txtCOA(2).Enabled = True
    txtCOA(2) = "$0.00"
    GetTextColor Me
  End With
  
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
  On Error GoTo DeleteErr
  With ADOprimaryrs
    If .RecordCount = 0 Then Exit Sub   ' no records!!
    If .EditMode = 0 Then
       If txtCOA(2) <> 0 Then
            MsgBox "Can't delete active COA", vbInformation, "Information"
            Exit Sub
       Else
        .Delete
       End If
        .MoveNext
        If .RecordCount = 0 Then     'no more records
            cmdUpdate.Enabled = False
            cmdDelete.Enabled = False
            cmdRefresh.Enabled = False
            BalBut.Enabled = False
            HistBut.Enabled = False
            .Requery
            Exit Sub
        ElseIf .EOF Then
            .MoveLast
        End If
        If Not (.BOF And .EOF) Then Bookmark = .Bookmark
    Else
        MsgBox "Must update or refresh record before deleting.", vbCritical, _
            "Delete Error."
    End If
  End With
  Exit Sub
DeleteErr:
  MsgBox Err.Description
End Sub

Private Sub cmdRefresh_Click()
    Set grdDataGrid.DataSource = Nothing
    ADOprimaryrs.Close
    Set ADOprimaryrs = Nothing
    OpenDB
End Sub

Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr
  
  With ADOprimaryrs
    If .RecordCount = 0 Then Exit Sub ' no records to update
    .Update
    If mbAddNewFlag Then
        Dim srch As String
        srch = .Fields("GL COA Account No").Value
        .Requery
        .Find "[GL COA Account No] = '" & srch & "'"
    End If
    'MsgBox .EOF
    mbAddNewFlag = False
  End With
  SetButtons True, 1
  Exit Sub
UpdateErr:
  'MsgBox Err.Number & "  " & Err.Description
  MsgBox Err.Description
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdFirst_Click()
  On Error GoTo GoFirstError

  ADOprimaryrs.MoveFirst
  'mbDataChanged = False

  Exit Sub

GoFirstError:
  MsgBox Err.Description
End Sub

Private Sub cmdLast_Click()
  On Error GoTo GoLastError

  ADOprimaryrs.MoveLast
  'mbDataChanged = False

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
  'mbDataChanged = False

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
  'mbDataChanged = False

  Exit Sub

GoPrevError:
  MsgBox Err.Description
End Sub

Private Sub SetButtons(bVal As Boolean, Optional DoItorNot As Integer)
  'cmdAdd.Enabled = bVal
  
  If DoItorNot = 1 Then
       cmdUpdate.Enabled = True
       cmdCancel.Visible = Not bVal
       cmdAdd.Enabled = True
       cmdAdd.Visible = True
  ElseIf DoItorNot = 3 Then
       cmdUpdate.Enabled = False
       cmdAdd.Enabled = False
  Else
       cmdUpdate.Enabled = True
       cmdCancel.Visible = Not bVal
    If cmdCancel.Visible = True Then
       cmdCancel.Left = cmdAdd.Left
       cmdAdd.Visible = False
    Else
       ADOprimaryrs(0) = "AAA"
       ADOprimaryrs.Update
       ADOprimaryrs.Delete
       cmdAdd.Visible = True
       'cmdAdd.ZOrder 0
    End If
  End If
  cmdCOA.Enabled = bVal
  cmdDelete.Enabled = bVal
  'cmdClose.Enabled = bVal
  cmdRefresh.Enabled = bVal
  cmdNext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdLast.Enabled = bVal
  cmdPrevious.Enabled = bVal
  cmdSearch.Enabled = bVal
  BalBut.Enabled = bVal
  HistBut.Enabled = bVal
If mbAddNewFlag = False Then
  Frame1.Enabled = bVal
  GetTextColor Me
End If
End Sub

Private Sub grdDataGrid_AfterUpdate()
  'reenable buttons
  cmdAdd.Enabled = True
  BalBut.Enabled = True
  HistBut.Enabled = True
  cmdDelete.Enabled = True
  cmdRefresh.Enabled = True
End Sub

Private Sub grdDataGrid_BeforeDelete(Cancel As Integer)
    If mbAddNewFlag = True Then Exit Sub
    Dim DeleteCration As Integer
    
    DeleteCration = MsgBox("Attempting to delete the data. " & vbCr & "Are you sure?", vbYesNo, "Deleting Confirmation")
    If DeleteCration = vbNo Then Cancel = 1
End Sub

Private Sub grdDataGrid_BeforeInsert(Cancel As Integer)
  'enable/disable buttons
    'cmdAdd.Enabled = False
    'cmdUpdate.Enabled = True
    'cmdDelete.Enabled = False
    'cmdRefresh.Enabled = False
    'BalBut.Enabled = False
    'HistBut.Enabled = False
    SetButtons False
End Sub

Private Sub grdDataGrid_Error(ByVal DataError As Integer, Response As Integer)
    If DataGridKnownError(DataError) Then
        Response = 0
    End If
End Sub

Private Sub grdDataGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  If Not ADOprimaryrs.BOF Or Not ADOprimaryrs.EOF Then
    If mbAddNewFlag = False Then
        If grdDataGrid.Row > 0 Then
            Dim TextCOA As TextBox
            If grdDataGrid.Columns(3) <> "$0.00" Then
                For Each TextCOA In Me.txtCOA
                    TextCOA.Enabled = False
                Next
                cbCOA(1).Enabled = False
            Else
                For Each TextCOA In Me.txtCOA
                    TextCOA.Enabled = True
                Next
                cbCOA(1).Enabled = True
            End If
            GetTextColor Me
        End If
    End If
  End If
End Sub

Private Sub HistBut_Click()
    If IsNull(grdDataGrid.Columns(0).Value) Then
        MsgBox "Please select a valid Account Number", vbCritical, _
        "Error: Invalid Account Number"
        Exit Sub
    Else
        frm_GL_Account_History.OpenAccount grdDataGrid.Columns(0).Value
        frm_GL_Account_History.Show
    End If
End Sub

Private Sub grdDataGrid_HeadClick(ByVal ColIndex As Integer)
If ADOprimaryrs.RecordCount = 0 Then Exit Sub
    lblLabels(5) = grdDataGrid.Columns(ColIndex).Caption
    WhichField = grdDataGrid.Columns(ColIndex).DataField
    ADOprimaryrs.Close
    Set ADOprimaryrs = Nothing
    Set ADOprimaryrs = New ADODB.Recordset
    ADOprimaryrs.Open "select [GL COA Account No],[GL COA Account Name],[GL COA Asset Type],[GL COA Account Balance],[GL COA Balance Type] from [GL Chart Of Accounts] Order by [" & grdDataGrid.Columns(ColIndex).DataField & "]", db, adOpenKeyset, adLockOptimistic, adCmdText
    Set grdDataGrid.DataSource = ADOprimaryrs
SetDataSource
End Sub

Private Sub optBussType_Click(Index As Integer)
  Dim SQLstatement As String
  Dim Business As String
  Dim rsBuss As ADODB.Recordset

  ShowStatus True

  Select Case Index
  Case 0
    Business = "C Corp"
  Case 1
    Business = "S Corp"
  Case 2
    Business = "Partnership"
  End Select
  
  lblBussType = Business
  'Build SQL
  SQLstatement = "SELECT DISTINCT [GL COA Industry] FROM [GL Industry Chart Of Accounts]"
  
  'Debug.Print SQLstatement
  Set rsBuss = New ADODB.Recordset
  rsBuss.Open SQLstatement, db, adOpenKeyset, adLockOptimistic, adCmdText

  Dim i As Integer
  Dim iChar As Integer
  Dim strtemp As String
  
  lstBussType.Clear
  With rsBuss
    .MoveFirst
    i = 0
    Do While Not .EOF
      iChar = InStr(1, ![GL COA Industry], "- " & Business)
      If iChar <> 0 Then
        strtemp = Left(![GL COA Industry], iChar - 2)
        lstBussType.List(i) = strtemp
        i = i + 1
      End If
      .MoveNext
    Loop
  End With
  ShowStatus False
End Sub

Private Sub CreateBankCards()

  'On Error GoTo CreateBankCards_Error

  Dim rs As ADODB.Recordset
  Dim ID&

  Set rs = New ADODB.Recordset
  rs.Open "SELECT * FROM [GL Chart Of Accounts] where [GL COA Asset Type] = 'Cash'", db, adOpenKeyset, adLockOptimistic, adCmdText
  'On Error Resume Next
  
  Dim rsBank As ADODB.Recordset
  Set rsBank = New ADODB.Recordset
  rsBank.Open "[BANK Accounts]", db, adOpenKeyset, adLockOptimistic, adCmdTable
  
  rs.MoveFirst
    Do While Not rs.EOF
      ID& = rs("GL COA Account No")
      If rsBank.RecordCount > 0 Then
          rsBank.MoveFirst
      End If
      rsBank.Find "[BANK ACCT ID]='" & CStr(ID&) & "'"
      If rsBank.EOF Then
        'Create a bank card for this guy
        rsBank.AddNew
          rsBank("BANK ACCT ID") = CStr(ID&)
          rsBank("BANK ACCT Name") = rs("GL COA Account Name") & ""
          rsBank("BANK ACCT Number") = 0
          rsBank("BANK ACCT Next Check No") = 0
          rsBank("BANK ACCT GL Cash Acct") = CStr(ID&)
        rsBank.Update
      End If
      rs.MoveNext
    Loop

  Dim rsGLCOA As ADODB.Recordset
  Set rsGLCOA = New ADODB.Recordset
  rsGLCOA.Open "[GL Chart Of Accounts]", db, adOpenKeyset, adLockOptimistic, adCmdTable

  'Now delete bank cards that are no longer in COA
  'rsBank.Index = "PrimaryKey"
  'On Error Resume Next
  'Err = 0
  rsBank.MoveFirst
  'If Err = 0 Then
    Do While Not rsBank.EOF
      ID& = rsBank("BANK ACCT ID")
      'rsGLCOA.Index = "PrimaryKey"
      'rsGLCOA.Seek "=", ID&
      rsGLCOA.MoveFirst
      rsGLCOA.Find "[GL COA Account No]='" & ID& & "'"
      If rsGLCOA.EOF Then
        rsBank.Delete
      End If
      'Err = 0
      rsBank.MoveNext
      'If Err = 3021 Then Exit Do
    Loop
  'End If

rs.Close
rsBank.Close
rsGLCOA.Close
Set rs = Nothing
Set rsBank = Nothing
Set rsGLCOA = Nothing

Exit Sub

CreateBankError:
  MsgBox Error, , "Create Bank"
  Resume Next

CreateBankCards_Error:
  Call ErrorLog("Load Chart Of Accounts", "CreateBankCards", Now, Err.Number, Err.Description, True, db)
  Resume Next
  
End Sub

Private Sub txtCOA_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
Case 0
    keyResponse = CtrlValidate(KeyAscii, "0123456789")
    If keyResponse = True Then
    Else
       KeyAscii = 0
    End If
Case 2
    keyResponse = CtrlValidate(KeyAscii, "0123456789.")
    If keyResponse = True Then
    Else
       KeyAscii = 0
    End If
End Select
End Sub

Private Sub txtCOA_LostFocus(Index As Integer)
If Index = 2 And txtCOA(Index) = "" Then
    txtCOA(2) = "$0.00"
End If
End Sub
