VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_GL_Entry 
   Caption         =   "GL Entry"
   ClientHeight    =   5595
   ClientLeft      =   1950
   ClientTop       =   3030
   ClientWidth     =   10710
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5595
   ScaleWidth      =   10710
   Begin VB.PictureBox picStatBox 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   10710
      TabIndex        =   36
      Top             =   5295
      Width           =   10710
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Picture         =   "frm_GL_Entry.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Picture         =   "frm_GL_Entry.frx":0342
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Picture         =   "frm_GL_Entry.frx":0684
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Picture         =   "frm_GL_Entry.frx":09C6
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   41
         Top             =   0
         Width           =   3360
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
      ScaleWidth      =   10710
      TabIndex        =   29
      Top             =   4995
      Width           =   10710
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   300
         Left            =   5400
         TabIndex        =   30
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   300
         Left            =   4320
         TabIndex        =   31
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   300
         Left            =   3240
         TabIndex        =   32
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   300
         Left            =   2160
         TabIndex        =   33
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   300
         Left            =   1080
         TabIndex        =   34
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   300
         Left            =   0
         TabIndex        =   35
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.Frame frPrimary 
      Height          =   4455
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   10695
      Begin VB.CommandButton cmdPost 
         Caption         =   "&Post"
         Height          =   780
         Left            =   9480
         Picture         =   "frm_GL_Entry.frx":0D08
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton cmdReverse 
         Caption         =   "Re&verse"
         Height          =   780
         Left            =   9480
         Picture         =   "frm_GL_Entry.frx":114A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdLookupGL 
         Height          =   285
         Left            =   3000
         Picture         =   "frm_GL_Entry.frx":158C
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1440
         Width           =   375
      End
      Begin MSDataGridLib.DataGrid grdDataGrid 
         Height          =   2385
         Left            =   240
         TabIndex        =   4
         Top             =   1920
         Width           =   6240
         _ExtentX        =   11007
         _ExtentY        =   4207
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
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "GL TRANSD Account"
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
            DataField       =   "GL TRANSD Debit Amount"
            Caption         =   "Debit Amount"
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
         BeginProperty Column02 
            DataField       =   "GL TRANSD Credit Amount"
            Caption         =   "Credit Amount"
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
         BeginProperty Column03 
            DataField       =   "GL TRANSD Project"
            Caption         =   "Project"
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
               Object.Visible         =   -1  'True
               ColumnWidth     =   1244.976
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2115.213
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   2145.26
            EndProperty
            BeginProperty Column03 
               Object.Visible         =   0   'False
               ColumnWidth     =   1349.858
            EndProperty
         EndProperty
      End
      Begin VB.PictureBox picMajor 
         BorderStyle     =   0  'None
         Height          =   4095
         Left            =   120
         ScaleHeight     =   4095
         ScaleWidth      =   10455
         TabIndex        =   5
         Top             =   240
         Width           =   10455
         Begin VB.TextBox txtFields 
            DataField       =   "GL TRANS Document #"
            Height          =   285
            Index           =   2
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   1200
            Width           =   1095
         End
         Begin VB.TextBox txtFields 
            DataField       =   "GL TRANS Number"
            Height          =   285
            Index           =   1
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   21
            Top             =   1200
            Width           =   1455
         End
         Begin VB.TextBox txtFields 
            DataField       =   "GL TRANS Date"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "M/d/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   20
            Top             =   480
            Width           =   1095
         End
         Begin VB.TextBox txtFields 
            DataField       =   "GL TRANS Reference"
            Height          =   645
            Index           =   3
            Left            =   6480
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   19
            Top             =   3360
            Width           =   3975
         End
         Begin VB.TextBox txtFields 
            DataField       =   "GL TRANS Description"
            Height          =   1125
            Index           =   5
            Left            =   6480
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   18
            Top             =   1920
            Width           =   3975
         End
         Begin VB.Frame Frame1 
            Height          =   1575
            Left            =   3360
            TabIndex        =   7
            Top             =   0
            Width           =   5895
            Begin VB.ComboBox cbfields 
               DataField       =   "GL TRANS Type"
               Height          =   315
               Index           =   6
               ItemData        =   "frm_GL_Entry.frx":1896
               Left            =   960
               List            =   "frm_GL_Entry.frx":1898
               Style           =   1  'Simple Combo
               TabIndex        =   42
               Text            =   "cbfields"
               Top             =   240
               Width           =   1575
            End
            Begin VB.CheckBox chkFields 
               Alignment       =   1  'Right Justify
               Caption         =   "Recurring Entry::"
               DataField       =   "GL TRANS Recurring YN"
               Height          =   285
               Index           =   0
               Left            =   3960
               TabIndex        =   16
               Top             =   240
               Width           =   1575
            End
            Begin VB.PictureBox Picture1 
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   900
               Left            =   120
               ScaleHeight     =   900
               ScaleWidth      =   5535
               TabIndex        =   8
               Top             =   600
               Width           =   5535
               Begin VB.TextBox txtFields 
                  BeginProperty DataFormat 
                     Type            =   1
                     Format          =   """$""#,##0.00"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   2
                  EndProperty
                  Height          =   285
                  Index           =   4
                  Left            =   120
                  TabIndex        =   12
                  Top             =   600
                  Width           =   1575
               End
               Begin VB.TextBox txtFields 
                  BeginProperty DataFormat 
                     Type            =   1
                     Format          =   """$""#,##0.00"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   2
                  EndProperty
                  Height          =   285
                  Index           =   8
                  Left            =   3960
                  TabIndex        =   11
                  Top             =   600
                  Width           =   1575
               End
               Begin VB.TextBox txtFields 
                  BeginProperty DataFormat 
                     Type            =   1
                     Format          =   """$""#,##0.00"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   2
                  EndProperty
                  Height          =   285
                  Index           =   9
                  Left            =   2040
                  TabIndex        =   10
                  Top             =   600
                  Width           =   1575
               End
               Begin VB.CheckBox chkFields 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Posted YN:"
                  DataField       =   "GL TRANS Posted YN"
                  Height          =   285
                  Index           =   5
                  Left            =   120
                  TabIndex        =   9
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.Label lblfields 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "Balance:"
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Index           =   4
                  Left            =   120
                  TabIndex        =   15
                  Top             =   360
                  Width           =   1575
               End
               Begin VB.Label lblfields 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "Credit Total:"
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Index           =   8
                  Left            =   3960
                  TabIndex        =   14
                  Top             =   360
                  Width           =   1575
               End
               Begin VB.Label lblfields 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "Debit Total:"
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Index           =   9
                  Left            =   2040
                  TabIndex        =   13
                  Top             =   360
                  Width           =   1575
               End
            End
            Begin VB.Label lblLabels 
               Caption         =   "GL Type:"
               Height          =   255
               Index           =   6
               Left            =   240
               TabIndex        =   17
               Top             =   240
               Width           =   975
            End
         End
         Begin VB.CommandButton cmdDate 
            Height          =   285
            Index           =   20
            Left            =   1200
            Picture         =   "frm_GL_Entry.frx":189A
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   480
            Width           =   375
         End
         Begin VB.Label lblfields 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Date:"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label lblfields 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Document No:"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   26
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label lblfields 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Ext Document No:"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   1800
            TabIndex        =   25
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Reference:"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   6480
            TabIndex        =   24
            Top             =   3120
            Width           =   3975
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Description:"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   5
            Left            =   6480
            TabIndex        =   23
            Top             =   1680
            Width           =   3975
         End
      End
   End
   Begin VB.Label lbltop 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bank Reconciliation"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   28
      Top             =   120
      Width           =   7305
   End
End
Attribute VB_Name = "frm_GL_Entry"
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
'Dim grdOnAddNew As Boolean

Private Function CheckEmpty() As Boolean
 For Each Ctrl In Me.txtFields
    If Ctrl.Text = "" Then
      Select Case Ctrl.Index
      Case 1, 2
      Case Else
        MsgBox "There is an empty data in " & lblLabels(Ctrl.Index), vbInformation, "Empty Data"
        CheckEmpty = False
        Exit Function
      End Select
    End If
 Next
 If cbfields(6).Text = "" Then
    MsgBox "There is an empty data in " & lblLabels(6).Caption, vbInformation, "Empty Data"
    cbfields(6).Text = cbfields(6).List(1)
    CheckEmpty = False
    Exit Function
 End If
 CheckEmpty = True
End Function

Private Sub OpenDB(SQLstatement As String)
  NewLoad = True
  Set ADOprimaryrs = New ADODB.Recordset
  'adoPrimaryRS.Open "SHAPE {select [GL TRANS Number],[GL TRANS Date],[GL TRANS Document #],[GL TRANS Type],[GL TRANS Posted YN],[GL TRANS Recurring YN],[GL TRANS Description],[GL TRANS Reference] FROM [GL Transaction] ORDER BY [GL TRANS Number]} AS ParentCMD APPEND ({select [GL TRANSD Number], [GL TRANSD Account], [GL TRANSD Debit Amount], [GL TRANSD Credit Amount] FROM [GL Transaction Detail] Order by [GL TRANSD Number] } AS ChildCMD RELATE [GL TRANS Number] TO [GL TRANSD Number]) AS ChildCMD", db, adOpenStatic, adLockOptimistic
  ADOprimaryrs.Open SQLstatement, db, adOpenStatic, adLockOptimistic

  Dim Ctrl As Control
  For Each Ctrl In Me.Controls
    If TypeOf Ctrl Is TextBox Or TypeOf Ctrl Is CheckBox Or TypeOf Ctrl Is ComboBox Then
        Set Ctrl.DataSource = ADOprimaryrs
        If TypeOf Ctrl Is TextBox And Ctrl.DataField <> "" Then
           If ADOprimaryrs("" & Ctrl.DataField & "").Type = 202 Then Ctrl.MaxLength = ADOprimaryrs("" & Ctrl.DataField & "").DefinedSize
        End If
    End If
  Next

  If CheckNewDB(ADOprimaryrs, "Order Entry") = True Then
    cmdAdd_Click
  Else
    Set grdDataGrid.DataSource = ADOprimaryrs("ChildCMD").UnderlyingValue
    
    grdDataGrid.Columns(0).Button = True
    grdDataGrid.Columns(3).Button = True
    
    NewLoad = False
    CalcTotals
  End If
  
  
  NewLoad = False

End Sub

'Private Sub cbfields_LostFocus(Index As Integer)
'    If CheckCombo(cbfields(Index)) Then
'        MsgBox "Attempting to alter the selection! Please make your selection", vbCritical, "Error"
'        cbfields(Index).Text = cbfields(Index).List(1)
'        cbfields(Index).SetFocus
'    End If
'End Sub

Private Sub cmdDate_Click(Index As Integer)
    Menu_Calendar.WhoCallMe True, 1302
    'Menu_Calendar.Show vbModal
End Sub
Private Function Datavalidate%()

  'On Error GoTo DataValidate_Error
  
  'Verify Date
  If txtFields(0) = "" Then
    MsgBox "Enter a valid date!", , "Unable To Post"
    Datavalidate% = False
    Exit Function
  End If

  If Len(txtFields(5)) = 0 Then
    MsgBox "Enter a reason!", , "Unable To Post"
    Datavalidate% = False
    Exit Function
  End If
  If txtFields(9) = 0 Then
    MsgBox "No Item to post!", , "Unable To Post"
    Datavalidate% = False
    Exit Function
  End If
  'Verify post date
  Dim PeriodToPost%
  Dim PeriodClosed%

  Call VerifyPeriod(txtFields(0), PeriodToPost%, PeriodClosed%)
  If PeriodClosed% = True Then
    MsgBox "Unable to post to a closed period!", , "Unable To Post"
    Datavalidate% = False
    Exit Function
  End If

  Datavalidate% = True

  Exit Function
DataValidate_Error:
  Call ErrorLog("Inventory Adjustment", "DataValidate", Now, Err.Number, Err.Description, True, db)
  Resume Next

End Function


Private Sub cmdLookupGL_Click()
   AllLookup.ToWhichRecord ADOprimaryrs, "GL Entry", "ID//Type//Date//Reference"
   'AllLookup.Show vbModal
End Sub

Private Sub cmdPost_Click()

'  On Error GoTo cmdPost_Click_Error

  Dim Success%

'  Forms![GL Entry].[GL Entry Detail].Form.Refresh

  'Force record save
'  DoCmd.RunMacro "Save Record"
  If CheckEmpty = False Then Exit Sub
  If Datavalidate() = True Then
    
    db.BeginTrans

    'On Error GoTo PostError

    ShowStatus True

    'Post each account
    Dim rsDetail As ADODB.Recordset
    Dim DebitAmount@
    Dim CreditAmount@
    Dim Account$

    Set rsDetail = New ADODB.Recordset
    rsDetail.Open "SELECT * FROM [GL Transaction Detail] where [GL TRANSD Number] = " & txtFields(1), db, adOpenStatic, adLockOptimistic, adCmdText
    'On Error Resume Next
    If rsDetail.RecordCount = 0 Then
      MsgBox "No detail to post!", , "Unable To Post"
      ShowStatus False
      Exit Sub
    End If
    rsDetail.MoveFirst
    Do While Not rsDetail.EOF
      Account$ = rsDetail("GL TRANSD Account")
      DebitAmount@ = IIf(IsNull(rsDetail("GL TRANSD Debit Amount")), 0, rsDetail("GL TRANSD Debit Amount"))
      CreditAmount@ = IIf(IsNull(rsDetail("GL TRANSD Credit Amount")), 0, rsDetail("GL TRANSD Credit Amount"))
      Success% = PostCOA(Account$, txtFields(0), DebitAmount@, CreditAmount@)
      rsDetail.MoveNext
    Loop
    rsDetail.Close
    Set rsDetail = Nothing
    ShowStatus False

    chkFields(5).Value = 1
    cbfields(6).Text = "GL Entry"
    ADOprimaryrs![GL TRANS Source] = "GL " & txtFields(2)
    ADOprimaryrs.Update
  
    db.CommitTrans

    MsgBox "Transaction Posted."
    
    cmdPost.Enabled = False

    'DoCmd.GoToRecord A_FORM, "GL Entry", A_NEWREC
    'DoCmd.GoToControl "GL TRANS Description"

  End If

  Exit Sub

PostError:
  db.RollbackTrans
  MsgBox "An error occurred posting the transaction!", , "Error"
  Exit Sub

cmdPost_Click_Error:
  Call ErrorLog("GL Entry", "cmdPost_Click", Now, Err.Number, Err.Description, True, db)
  Resume Next

End Sub


Private Sub INV_ITEM()
   AllLookup.GetWhichTable 1420, "SELECT [INV ITEM Id], [INV ITEM Description]," & _
   "[INV ITEM Price],[INV ITEM Inventory Account], [INV ITEM Qty On Hand], " & _
   "[INV ITEM Qty On Order], [INV ITEM Last Cost],[INV ITEM Average Cost] FROM [INV Items] " & _
   "WHERE [INV ITEM Type] = 'Assembly' ", "Product", _
   "Item ID//Item Description//Price//Inventory Account//Qty On Hand//Qty On Order", db
   'AllLookup.Show vbModal

End Sub

Private Sub COA_grdDataGrid_Butt()
   AllLookup.GetWhichTable 1450, "Select [GL COA Account No],[GL COA Account Name]," & _
   "[GL COA Asset Type] From [GL Chart Of Accounts] ", "GL Accounts", _
   "Account No//Account Type//Account Type", db
   'AllLookup.Show vbModal
   
End Sub

Private Sub Proj_Projects()
   AllLookup.GetWhichTable 1453, "Select [PROJ ID],[PROJ Name]," & _
   "[PROJ Description] From [PROJ Projects] ", "Project", _
   "Project ID//Project Name//Description", db
   'AllLookup.Show vbModal
   
End Sub

Private Sub cmdReverse_Click()

'  On Error GoTo cmdReverse_Click_Error

  'Copy this entry and detail to new records
  
'  If IsNull(Me![GL TRANS Document #]) Then Exit Sub
'  If Len(Me![GL TRANS Document #]) = 0 Then Exit Sub

  Dim Success%
  Success% = ReverseGLEntry(CLng(txtFields(1)), True)
  If Success% = 1 Then
    'Reverse was cancelled
    Exit Sub
  End If
  If Success% = False Then
    MsgBox "Reverse failed!", , "Error"
    Exit Sub
  End If

  'Refresh the recordset
  'Me.Requery

  'Goto last record in table
  'DoCmd.GoToRecord A_FORM, "GL Entry", A_LAST

  Exit Sub
cmdReverse_Click_Error:
  Call ErrorLog("GL Entry", "cmdReverse_Click", Now, Err.Number, Err.Description, True, db)
  Resume Next

End Sub


Private Sub Form_Load()
On Error GoTo FormErr
ShowStatus True
  Set db = New ADODB.Connection
  db.CursorLocation = adUseClient
  db.Provider = "MSDataShape"
  db.Open "Data " & gblADOProvider
  
  OpenDB "SHAPE {select [GL TRANS Document #],[GL TRANS Type],[GL TRANS Date],[GL TRANS Reference],[GL TRANS Number],[GL TRANS Posted YN],[GL TRANS Recurring YN],[GL TRANS Description],[GL TRANS Source] FROM [GL Transaction] ORDER BY [GL TRANS Number]} AS ParentCMD APPEND ({select [GL TRANSD Number], [GL TRANSD Account], [GL TRANSD Debit Amount], [GL TRANSD Credit Amount] FROM [GL Transaction Detail]} AS ChildCMD RELATE [GL TRANS Number] TO [GL TRANSD Number]) AS ChildCMD"
  
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
  
  Me.Width = 10830
  Me.Height = 6000
  
SkipResize:
  frPrimary.Left = (Me.ScaleWidth - frPrimary.Width) / 2
  Picture1.Left = frPrimary.Left
  lbltop.Width = frPrimary.Width
  lbltop.Left = frPrimary.Left
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
    ShowStatus False
    If UnloadForm(ADOprimaryrs) = 0 Then
        db.Close
        Set db = Nothing
    Else
        Cancel = 1
    End If
Exit Sub
FormErr:
  MsgBox Err.Description
  ShowStatus False
End Sub

Private Sub adoPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This will display the current record position for this recordset
  If ADOprimaryrs.BOF Or ADOprimaryrs.EOF Then Exit Sub
  If ADOprimaryrs![GL TRANS Posted YN] = True Then
    picMajor.Enabled = False
    cmdPost.Enabled = False
    cmdReverse.Enabled = True
    grdDataGrid.Enabled = False
  Else
    CalcTotals
    picMajor.Enabled = True
    cmdPost.Enabled = True
    cmdReverse.Enabled = False
    grdDataGrid.Enabled = True
  End If
  'If NewLoad = False Then
  '  If IsNull(adoPrimaryRS![INV PRO Notes]) Then
  '    txtFields(5).Text = Now
  '  End If
  '  If IsNull(adoPrimaryRS![INV PRO Memo]) Then
  '      txtFields(3) = txtFields(3) & "-Created by " & AppLoginName
  '  End If
  'End If
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
ShowStatus True
If cmdAdd.Caption = "&Save" Then
     If Not CheckEmpty Then
        ShowStatus False
        Exit Sub
     End If
     With ADOprimaryrs
         .UpdateBatch adAffectAll
         .MovePrevious
         grdDataGrid.HoldFields
         grdDataGrid.ReBind
         grdDataGrid.Refresh
         NewLoad = False
         cmdRefresh_Click
         .MoveLast
         txtFields(2) = "GL " & ![GL TRANS Number] + 1000
         .Update
         'txtFields(3).SetFocus
     End With
     cmdAdd.Caption = "&Add"
     mbAddNewFlag = False
     SetButtons True
     cmdPost.Enabled = True
     cmdReverse.Enabled = False
Else
  With ADOprimaryrs
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
     NewLoad = True
     '.MoveLast
'     Text1 = ![INV ADJ Ext Document No]
     .AddNew
'     txtFields(2) = Val(Text1) + 1
     cbfields(6).Enabled = True
     cbfields(6).Text = cbfields(6).List(1)
     lblStatus.Caption = "Add record"
     mbAddNewFlag = True
     cmdPost.Enabled = False
     cmdReverse.Enabled = True
     txtFields(3) = "-Created by " & AppLoginName
     txtFields(8) = "$0.00"
     txtFields(9) = "$0.00"
     txtFields(4) = "$0.00"
     SetButtons False
  End With
  cmdAdd.Caption = "&Save"
End If
  ShowStatus False
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
     If Not DataDelete(ADOprimaryrs, Me, True) Then
     End If
End Sub

Private Sub cmdRefresh_Click()
    RefreshButton ADOprimaryrs, grdDataGrid
    grdDataGrid.Columns(0).Button = True
    grdDataGrid.Columns(3).Button = True
End Sub

Private Sub cmdEdit_Click()
  'On Error GoTo EditErr

  lblStatus.Caption = "Edit record"
  mbEditFlag = True
  SetButtons True
  Exit Sub

EditErr:
  MsgBox Err.Description
End Sub
Private Sub cmdCancel_Click()
  ShowStatus True
  SetButtons True
  cmdAdd.Caption = "&Add"
  cmdCancel.Visible = False
  mbEditFlag = False
  mbAddNewFlag = False
  ADOprimaryrs.CancelUpdate
  NewLoad = False
  cmdPost.Enabled = True
  cmdReverse.Enabled = False
  If ADOprimaryrs.RecordCount > 0 Then
    ADOprimaryrs.MoveLast
  Else
    MsgBox "No data to publish. Exiting " & Me.Caption
    Unload Me
    Exit Sub
  End If
  ADOprimaryrs.Resync adAffectCurrent
  If mvBookMark > 0 Then
    ADOprimaryrs.Bookmark = mvBookMark
  Else
    ADOprimaryrs.MoveFirst
  End If
  mbDataChanged = False
  ShowStatus False
End Sub

Private Sub cmdUpdate_Click()
'Dim FlagStatus As Boolean
    
  'FlagStatus = False

  Call UpdateButton(ADOprimaryrs, mbAddNewFlag)
  
  'mbEditFlag = Not FlagStatus
  
  'SetButtons FlagStatus
  'mbDataChanged = Not FlagStatus
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
End Sub

Private Sub grdDataGrid_AfterColUpdate(ByVal ColIndex As Integer)
  If grdOnAddNew = True Then
     grdOnAddNew = False
  End If

End Sub

Private Sub grdDataGrid_BeforeDelete(Cancel As Integer)
    Dim DeleteCration As Integer
    
    DeleteCration = MsgBox("Attempting to delete the data. " & vbCr & "Are you sure?", vbYesNo, "Deleting Confirmation")
    If DeleteCration = vbNo Then Cancel = 1
End Sub

Private Sub grdDataGrid_ButtonClick(ByVal ColIndex As Integer)
On Error GoTo Error_ButtClick
If mbAddNewFlag = True Then Exit Sub
If grdDataGrid.Columns(0) <> "" Then grdOnAddNew = False
Select Case ColIndex
Case 0   'Get the type of account for the selected row
    COA_grdDataGrid_Butt
Case 3   'Select the project that have been working on
    Proj_Projects
End Select

If grdOnAddNew = True And grdDataGrid.Columns(0) <> "" Then NewgrdDatagrid
grdDataGrid_AfterColEdit 0
Exit Sub
Error_ButtClick:
    MsgBox "Please click the Table box before clicking the button"
End Sub
Private Sub grdDataGrid_AfterColEdit(ByVal ColIndex As Integer)
'If NewLoad = False Then
'  grdDataGrid.Columns(4).Text = grdDataGrid.Columns(2).Value * grdDataGrid.Columns(3).Value
'End If
    If grdDataGrid.Row = -1 Or grdDataGrid.Columns(0) = "" Then Exit Sub
      SendKeys ("{ENTER}")
  If grdDataGrid.Row > 0 Then
      SendKeys ("{up}")
      SendKeys ("{down}")
  ElseIf grdDataGrid.Row = 0 Then
      SendKeys ("{down}")
      SendKeys ("{up}")
  End If
End Sub

Private Sub CalcTotals()
If NewLoad = True Then Exit Sub
Dim TempDebit As Currency
Dim TempCredit As Currency
  Dim CalcTable  As ADODB.Recordset
  Set CalcTable = New ADODB.Recordset
  TempDebit = 0
  TempCredit = 0
  'balance=Sum([GL TRANSD Debit Amount])-Sum([GL TRANSD Credit Amount])
  'debit total=Sum([GL TRANSD Debit Amount])
  'credit total=Sum([GL TRANSD Credit Amount])
  'Debug.Print "SELECT [GL TRANSD Debit Amount],[GL TRANSD Credit Amount] FROM [GL Transaction Detail] WHERE [GL TRANSD Number]=" & adoPrimaryRS![GL TRANS Number]
  CalcTable.Open "SELECT [GL TRANSD Debit Amount],[GL TRANSD Credit Amount] FROM [GL Transaction Detail] WHERE [GL TRANSD Number]=" & ADOprimaryrs![GL TRANS Number], db, adOpenStatic, adLockOptimistic, adCmdText
  With CalcTable
    cbfields(6).Enabled = True
    If .RecordCount = 0 Then
      txtFields(8) = "$0.00"
      txtFields(9) = "$0.00"
      txtFields(4) = "$0.00"
      Exit Sub
    End If
    cbfields(6).Enabled = False
    .MoveFirst
    Do While Not .EOF
       TempDebit = TempDebit + ![GL TRANSD Debit Amount]
       TempCredit = TempCredit + ![GL TRANSD Credit Amount]
       .MoveNext
    Loop
  .Close
  End With
  Set CalcTable = Nothing
  txtFields(8) = FormatCurr(TempCredit)
  txtFields(9) = FormatCurr(TempDebit)
  txtFields(4) = FormatCurr(TempDebit - TempCredit)
End Sub

Private Sub NewgrdDatagrid()
    NewLoad = True
    NewRowForDataGrid ADOprimaryrs, grdDataGrid, "GL TRANS Description", txtFields(5).Text
    grdOnAddNew = False
    NewLoad = False '     NewLoad = True
'     With ADOprimaryrs
'        If .BOF = False Or .EOF = False Then
'           mvBookMark = .Bookmark
'        End If
'        ADOprimaryrs("GL TRANS Description") = txtFields(5) & ""
        'txtFields_LostFocus 3
        'txtFields_LostFocus 5
'        ADOprimaryrs.Update
'        Set grdDataGrid.DataSource = Nothing
            'adoPrimaryRS.UpdateBatch adAffectAll
'            ADOprimaryrs.Requery
'        Set grdDataGrid.DataSource = ADOprimaryrs("ChildCMD").UnderlyingValue
'        If mvBookMark > 0 Then
'           ADOprimaryrs.Bookmark = mvBookMark
'        Else
'           ADOprimaryrs.MoveFirst
'        End If
'    End With
    'Debug.Print mvBookMark
'    grdOnAddNew = False
'    NewLoad = False
End Sub

Private Sub grdDataGrid_Error(ByVal DataError As Integer, Response As Integer)
    If DataGridKnownError(DataError) Then
        Response = 0
    End If
End Sub

Private Sub grdDataGrid_GotFocus()
    If mbAddNewFlag = True Then
        cmdAdd.SetFocus
        CreateOrder = MsgBox("This Request will save the data to the database? Are sure to continue", vbYesNo, "Save Quote")
        If CreateOrder = vbNo Then Exit Sub
        cmdAdd_Click
    End If
End Sub

Private Sub grdDataGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If NewLoad = True Then Exit Sub
    If ADOprimaryrs.BOF Or ADOprimaryrs.EOF Then Exit Sub
        If grdDataGrid.col > 0 And grdDataGrid.Row > -1 Then
            If grdDataGrid.Columns(0) = "" Then
                MsgBox "You must select Item ID first before continue", vbInformation, "Error Selection"
                GoTo Damn_Attempt
            End If
        End If
CalcTotals
Select Case grdDataGrid.col
  Case 1
     grdDataGrid.AllowUpdate = True
  Case 2
     grdDataGrid.AllowUpdate = True
  Case Else
     grdDataGrid.AllowUpdate = False
  End Select
Exit Sub
Damn_Attempt:
     grdDataGrid.AllowUpdate = False
     grdDataGrid.col = 0
exit_sub:

End Sub
Private Sub grdDataGrid_OnAddNew()
    grdOnAddNew = True
End Sub


'Private Sub optSort_Click(Index As Integer)
'Select Case Index
'Case 0
'    OpenDb "SHAPE {select [GL TRANS Number],[GL TRANS Date],[GL TRANS Document #],[GL TRANS Type],[GL TRANS Posted YN],[GL TRANS Recurring YN],[GL TRANS Description],[GL TRANS Reference] FROM [GL Transaction] WHERE [GL TRANS Posted YN]=TRUE} AS ParentCMD APPEND ({select [GL TRANSD Number], [GL TRANSD Account], [GL TRANSD Debit Amount], [GL TRANSD Credit Amount] FROM [GL Transaction Detail]} AS ChildCMD RELATE [GL TRANS Number] TO [GL TRANSD Number]) AS ChildCMD"
'Case 1
'    OpenDb "SHAPE {select [GL TRANS Number],[GL TRANS Date],[GL TRANS Document #],[GL TRANS Type],[GL TRANS Posted YN],[GL TRANS Recurring YN],[GL TRANS Description],[GL TRANS Reference] FROM [GL Transaction] WHERE [GL TRANS Posted YN]=FALSE} AS ParentCMD APPEND ({select [GL TRANSD Number], [GL TRANSD Account], [GL TRANSD Debit Amount], [GL TRANSD Credit Amount] FROM [GL Transaction Detail]} AS ChildCMD RELATE [GL TRANS Number] TO [GL TRANSD Number]) AS ChildCMD"
'End Select
'End Sub

Private Sub txtFields_LostFocus(Index As Integer)
If NewLoad = True Then Exit Sub
Select Case Index
Case 0
    txtFields(0) = FormatDate(Now)
Case 3
    If InStr(1, txtFields(Index), "-Created by " & AppLoginName, vbTextCompare) Then Exit Sub
    txtFields(Index) = txtFields(Index) & "-Created by " & AppLoginName
Case 5
    If Len(Trim(txtFields(Index))) = 0 Then
       MsgBox "Please enter the Notes for the Production", vbCritical, "Error"
       txtFields(5).SetFocus
    End If
End Select

End Sub


