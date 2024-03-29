VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_INV_Production 
   Caption         =   "Inventory Production"
   ClientHeight    =   6510
   ClientLeft      =   1950
   ClientTop       =   3030
   ClientWidth     =   12270
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6510
   ScaleWidth      =   12270
   Begin VB.Frame frPrimary 
      Height          =   6015
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   12255
      Begin MSDataGridLib.DataGrid grdDataGrid 
         Height          =   3105
         Left            =   120
         TabIndex        =   41
         Top             =   2760
         Width           =   8760
         _ExtentX        =   15452
         _ExtentY        =   5477
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
         Caption         =   "Inventory Production"
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "INV PROD Item ID"
            Caption         =   "Item ID"
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
            DataField       =   "INV PROD Description"
            Caption         =   "Item Description"
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
            DataField       =   "INV PROD Quantity"
            Caption         =   "Prod Qty."
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
            DataField       =   "INV PROD Cost"
            Caption         =   "Prod. Cost"
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
         BeginProperty Column04 
            DataField       =   "INV PROD Ext Cost"
            Caption         =   "Ext. Cost"
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
         BeginProperty Column05 
            DataField       =   "INV PROD Project"
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
               ColumnWidth     =   1305.071
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3000.189
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1305.071
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1305.071
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1305.071
            EndProperty
            BeginProperty Column05 
               Object.Visible         =   0   'False
               ColumnWidth     =   1319.811
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame2 
         Height          =   5775
         Left            =   9000
         TabIndex        =   9
         Top             =   120
         Width           =   3135
         Begin VB.PictureBox picButtons 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   660
            Left            =   120
            ScaleHeight     =   660
            ScaleWidth      =   2955
            TabIndex        =   34
            Top             =   4680
            Width           =   2955
            Begin VB.CommandButton cmdCancel 
               Caption         =   "&Cancel"
               Height          =   300
               Left            =   1920
               TabIndex        =   35
               Top             =   360
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.CommandButton cmdDelete 
               Caption         =   "&Delete"
               Height          =   300
               Left            =   1920
               TabIndex        =   38
               Top             =   0
               Width           =   975
            End
            Begin VB.CommandButton cmdUpdate 
               Caption         =   "&Update"
               Height          =   300
               Left            =   960
               TabIndex        =   39
               Top             =   0
               Width           =   975
            End
            Begin VB.CommandButton cmdClose 
               Caption         =   "&Close"
               Height          =   300
               Left            =   960
               TabIndex        =   36
               Top             =   360
               Width           =   975
            End
            Begin VB.CommandButton cmdAdd 
               Caption         =   "&Add"
               Height          =   300
               Left            =   0
               TabIndex        =   40
               Top             =   0
               Width           =   975
            End
            Begin VB.CommandButton cmdRefresh 
               Caption         =   "&Refresh"
               Height          =   300
               Left            =   0
               TabIndex        =   37
               Top             =   360
               Width           =   975
            End
         End
         Begin VB.CommandButton cmdPost 
            Caption         =   "&Post"
            Height          =   780
            Left            =   2160
            Picture         =   "frm_INV_Production.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   3720
            Width           =   855
         End
         Begin VB.PictureBox picStatBox 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   120
            ScaleHeight     =   300
            ScaleWidth      =   2955
            TabIndex        =   27
            Top             =   5400
            Width           =   2955
            Begin VB.CommandButton cmdFirst 
               Height          =   300
               Left            =   0
               Picture         =   "frm_INV_Production.frx":030A
               Style           =   1  'Graphical
               TabIndex        =   31
               Top             =   0
               UseMaskColor    =   -1  'True
               Width           =   345
            End
            Begin VB.CommandButton cmdPrevious 
               Height          =   300
               Left            =   345
               Picture         =   "frm_INV_Production.frx":064C
               Style           =   1  'Graphical
               TabIndex        =   30
               Top             =   0
               UseMaskColor    =   -1  'True
               Width           =   345
            End
            Begin VB.CommandButton cmdLast 
               Height          =   300
               Left            =   2540
               Picture         =   "frm_INV_Production.frx":098E
               Style           =   1  'Graphical
               TabIndex        =   28
               Top             =   0
               UseMaskColor    =   -1  'True
               Width           =   345
            End
            Begin VB.CommandButton cmdNext 
               Height          =   300
               Left            =   2200
               Picture         =   "frm_INV_Production.frx":0CD0
               Style           =   1  'Graphical
               TabIndex        =   29
               Top             =   0
               UseMaskColor    =   -1  'True
               Width           =   345
            End
            Begin VB.Label lblStatus 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   690
               TabIndex        =   32
               Top             =   0
               Width           =   1520
            End
         End
         Begin VB.TextBox txtFields 
            DataField       =   "INV PRO Date"
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
            Index           =   0
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   23
            Top             =   3600
            Width           =   1215
         End
         Begin VB.TextBox txtFields 
            DataField       =   "INV PRO Ext Document No"
            Height          =   285
            Index           =   2
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   21
            Top             =   4200
            Width           =   1575
         End
         Begin VB.CommandButton cmdDate 
            Height          =   285
            Index           =   20
            Left            =   1440
            Picture         =   "frm_INV_Production.frx":1012
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   3600
            Width           =   375
         End
         Begin VB.Frame Frame1 
            Height          =   2535
            Left            =   120
            TabIndex        =   11
            Top             =   720
            Width           =   2895
            Begin VB.ComboBox cbfields 
               DataField       =   "INV PRO Type"
               Height          =   315
               Index           =   6
               ItemData        =   "frm_INV_Production.frx":131C
               Left            =   1440
               List            =   "frm_INV_Production.frx":1326
               TabIndex        =   12
               Text            =   "cbfields"
               Top             =   360
               Width           =   1335
            End
            Begin VB.PictureBox Picture1 
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   1425
               Left            =   120
               ScaleHeight     =   1425
               ScaleWidth      =   2655
               TabIndex        =   13
               Top             =   960
               Width           =   2655
               Begin VB.TextBox txtFields 
                  BeginProperty DataFormat 
                     Type            =   1
                     Format          =   """$""#,##0.00;(""$""#,##0.00)"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   2
                  EndProperty
                  Height          =   285
                  Index           =   8
                  Left            =   0
                  TabIndex        =   16
                  Top             =   1080
                  Width           =   1575
               End
               Begin VB.TextBox txtFields 
                  Height          =   285
                  Index           =   9
                  Left            =   0
                  TabIndex        =   15
                  Top             =   480
                  Width           =   1575
               End
               Begin VB.CheckBox chkFields 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Posted :"
                  DataField       =   "INV PRO Posted YN"
                  Height          =   285
                  Index           =   5
                  Left            =   1680
                  TabIndex        =   14
                  Top             =   240
                  Width           =   975
               End
               Begin VB.Label lblLabels 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "Total Cost:"
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Index           =   8
                  Left            =   0
                  TabIndex        =   18
                  Top             =   840
                  Width           =   1575
               End
               Begin VB.Label lblLabels 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "Total Production Qty"
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Index           =   9
                  Left            =   0
                  TabIndex        =   17
                  Top             =   240
                  Width           =   1575
               End
            End
            Begin VB.Label lblLabels 
               Caption         =   "Adjustment Type:"
               Height          =   255
               Index           =   6
               Left            =   120
               TabIndex        =   19
               Top             =   360
               Width           =   1335
            End
         End
         Begin VB.PictureBox picPosted 
            BorderStyle     =   0  'None
            Height          =   525
            Left            =   120
            Picture         =   "frm_INV_Production.frx":133E
            ScaleHeight     =   525
            ScaleWidth      =   2955
            TabIndex        =   10
            Top             =   240
            Width           =   2955
         End
         Begin VB.TextBox txtFields 
            DataField       =   "INV PRO Document No"
            Height          =   285
            Index           =   1
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   2880
            Width           =   1695
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Date:"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   26
            Top             =   3360
            Width           =   1575
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Document No:"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   24
            Top             =   3960
            Width           =   1575
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Document No:"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   1200
            TabIndex        =   25
            Top             =   2640
            Width           =   1695
         End
      End
      Begin VB.PictureBox picMajor 
         BorderStyle     =   0  'None
         Height          =   2535
         Left            =   120
         ScaleHeight     =   2535
         ScaleWidth      =   10935
         TabIndex        =   2
         Top             =   240
         Width           =   10935
         Begin VB.TextBox txtFields 
            DataField       =   "INV PRO Notes"
            Height          =   2205
            Index           =   3
            Left            =   4440
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   3
            Top             =   240
            Width           =   4335
         End
         Begin VB.TextBox txtFields 
            DataField       =   "INV PRO Memo"
            Height          =   2205
            Index           =   5
            Left            =   0
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            Top             =   240
            Width           =   4335
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Adjustment Reason:"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   5
            Left            =   0
            TabIndex        =   6
            Top             =   0
            Width           =   4335
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Notes:"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   4440
            TabIndex        =   5
            Top             =   0
            Width           =   4335
         End
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   8640
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Last Number"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   8640
         TabIndex        =   7
         Top             =   1440
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.Label lbltop 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Inventory Production"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   120
      Width           =   10185
   End
End
Attribute VB_Name = "frm_INV_Production"
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
Dim RSstatement As String
Dim DocType As String

Private Function CheckEmpty() As Boolean
 For Each Ctrl In Me.txtfields
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



Private Sub cbfields_LostFocus(Index As Integer)
    If CheckCombo(cbfields(Index)) Then
        MsgBox "Attempting to alter the selection! Please make your selection", vbCritical, "Error"
        cbfields(Index).Text = cbfields(Index).List(1)
        cbfields(Index).SetFocus
    End If
End Sub

Private Sub cmdDate_Click(Index As Integer)
    Menu_Calendar.WhoCallMe True, 1302
    'Menu_Calendar.Show vbModal
    txtfields(0).SetFocus
End Sub
Private Function Datavalidate%()

  'On Error GoTo DataValidate_Error
  
  'Verify Date
  If txtfields(0) = "" Then
    MsgBox "Enter a valid date!", , "Unable To Post"
    Datavalidate% = False
    Exit Function
  End If

  If Len(txtfields(5)) = 0 Then
    MsgBox "Enter a reason!", , "Unable To Post"
    Datavalidate% = False
    Exit Function
  End If
  If txtfields(9) = 0 Then
    MsgBox "No Item to post!", , "Unable To Post"
    Datavalidate% = False
    Exit Function
  End If
  'Verify post date
  Dim PeriodToPost%
  Dim PeriodClosed%

  Call VerifyPeriod(txtfields(0), PeriodToPost%, PeriodClosed%)
  If PeriodClosed% = True Then
    MsgBox "Unable to post to a closed period!", , "Unable To Post"
    Datavalidate% = False
    Exit Function
  End If

  Datavalidate% = True

  Exit Function
DataValidate_Error:
  Call ErrorLog("Inventory Production", "DataValidate", Now, Err.Number, Err.Description, True, db)
  Resume Next

End Function


Private Sub cmdPost_Click()

  'On Error GoTo cmdPost_Click_Error

  Dim Success%
  If CheckEmpty = False Then Exit Sub
  If Datavalidate() = False Then
    ShowStatus False
    Exit Sub
  End If
  cmdUpdate_Click
  
  'get a confirmation from user
  Dim CreateOrder As Integer
    
  CreateOrder = MsgBox("Posting the data. Are you sure?", vbYesNo, "Posting")
  If CreateOrder = vbNo Then Exit Sub

  ShowStatus True
  
  db.BeginTrans
    Success% = PostProduction(CLng(txtfields(1)), db)
    If Success% = False Then
      db.RollbackTrans
      MsgBox "Transaction NOT Posted."
    Else
      db.CommitTrans
      MsgBox "Transaction Posted."
      'chkFields(5).Value = 1
      ADOprimaryrs![INV PRO Posted YN] = True
      picPosted.Visible = True
      cmdPost.Enabled = False
      ADOprimaryrs.Update
      cmdPost.Enabled = False
    End If

  ShowStatus False
  
  Exit Sub
  
RecordLocked:
  db.RollbackTrans
  Exit Sub

UnableToPost:
  db.RollbackTrans
  Exit Sub

cmdPost_Click_Error:
  Call ErrorLog("Inventory Production", "cmdPost_Click", Now, Err.Number, Err.Description, True, db)
  Resume Next

End Sub

Private Sub INV_ITEM()
   AllLookup.GetWhichTable 1420, "SELECT [INV ITEM Id], [INV ITEM Description]," & _
   "[INV ITEM Price],[INV ITEM Inventory Account], [INV ITEM Qty On Hand], " & _
   "[INV ITEM Qty On Order], [INV ITEM Last Cost],[INV ITEM Average Cost] FROM [INV Items] " & _
   "WHERE [INV ITEM Type] = 'Assembly' AND [INV ITEM Inactive YN]=FALSE ", "Product", _
   "Item ID//Item Description//Price//Inventory Account//Qty On Hand//Qty On Order", db
   'AllLookup.Show vbModal 'Accounting Software

End Sub

Private Sub COA_grdDataGrid_Butt()
   AllLookup.GetWhichTable 1410, "Select [GL COA Account No],[GL COA Account Name]," & _
   "[GL COA Asset Type] From [GL Chart Of Accounts] ", "GL Accounts", _
   "Account No//Account Type//Account Type", db
   'AllLookup.Show vbModal
   
End Sub

Private Sub Proj_Projects()
   AllLookup.GetWhichTable 1403, "Select [PROJ ID],[PROJ Name]," & _
   "[PROJ Description] From [PROJ Projects] ", "Project", _
   "Project ID//Project Name//Description", db
   'AllLookup.Show vbModal
   
End Sub

Private Sub RedoNumbers()

  
End Sub


Private Sub Form_Load()
'On Error GoTo FormErr
ShowStatus True
  DocType = "TempINVPro"
  
  Set db = New ADODB.Connection
  db.CursorLocation = adUseClient
  db.Provider = "MSDataShape"
  db.Open "Data " & gblADOProvider
  
  RSstatement = "SHAPE {select [INV PRO Date],[INV PRO Document No],[INV PRO Ext Document No],[INV PRO Memo],[INV PRO Notes],[INV PRO Posted YN],[INV PRO Type] from [INV Production] ORDER BY [INV PRO Document No]} AS ParentCMD APPEND ({select [INV PROD ID],[INV PROD Document No],[INV PROD Item ID],[INV PROD Description],[INV PROD Quantity],[INV PROD Cost],[INV PROD Ext Cost],[INV PROD Project] from [INV Production Detail] Order by [INV PROD Item ID] } AS ChildCMD RELATE [INV PRO Document No] TO [INV PROD Document No]) AS ChildCMD"
  OpenDB RSstatement
  
  grdDataGrid.Columns(0).Button = True
  grdDataGrid.Columns(5).Button = True
  'grdDataGrid.Columns(6).Button = True
  
  grdDataGrid.AllowAddNew = True
  grdDataGrid.AllowDelete = True
  
  GetTextColor Me
  mbDataChanged = False
  ShowStatus False
  NewLoad = False
Exit Sub
FormErr:
  MsgBox Err.Description
  ShowStatus False
End Sub

Private Sub OpenDB(SQLstatement As String, Optional NewData As Boolean)
  NewLoad = True
  ShowStatus True

  Set ADOprimaryrs = New ADODB.Recordset
  ADOprimaryrs.Open SQLstatement, db, adOpenStatic, adLockOptimistic
  With ADOprimaryrs
    If NewData = True Then
        ADOprimaryrs.Find "[INV PRO Ext Document No]='" & DocType & AppLoginName & "'"
      If Not .EOF Then
        ADOprimaryrs![INV PRO Ext Document No] = AppLoginName & Format(Now, "MMdd") & Right(Format(![INV PRO Document No], "0000"), 4)
        'ADOprimaryrs![AR SALE Status] = "Open"
        ADOprimaryrs.Update
      Else
        .MoveFirst
      End If
    End If
  End With
  Dim Ctrl As Control
  For Each Ctrl In Me.Controls
    If TypeOf Ctrl Is TextBox Or TypeOf Ctrl Is CheckBox Or TypeOf Ctrl Is ComboBox Then
        Set Ctrl.DataSource = ADOprimaryrs
        If TypeOf Ctrl Is TextBox And Ctrl.DataField <> "" Then
           If ADOprimaryrs("" & Ctrl.DataField & "").Type = 202 Then Ctrl.MaxLength = ADOprimaryrs("" & Ctrl.DataField & "").DefinedSize
        End If
    End If
  Next

  If CheckNewDB(ADOprimaryrs, "Inventory Production") = True Then
    cmdAdd_Click
  Else
    Set grdDataGrid.DataSource = ADOprimaryrs("ChildCMD").UnderlyingValue
    CalcTotals
  End If
NewLoad = False
End Sub

Private Sub ClearDatasource()
 Dim Ctrl As Control
 For Each Ctrl In Me.Controls
    If TypeOf Ctrl Is TextBox Or TypeOf Ctrl Is CheckBox Or TypeOf Ctrl Is ComboBox Then
        If Ctrl.DataField <> "" Then
           Set Ctrl.DataSource = Nothing
        End If
    End If
 Next
    Set grdDataGrid.DataSource = Nothing
    ADOprimaryrs.CancelUpdate
    ADOprimaryrs.Close
    Set ADOprimaryrs = Nothing
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
  
  Me.Width = 12390
  Me.Height = 6915
SkipResize:
  frPrimary.Left = (Me.ScaleWidth - frPrimary.Width) / 2
  frPrimary.Top = (Me.ScaleHeight - frPrimary.Height) / 2 + 240
  lblTop.Left = frPrimary.Left
  lblTop.Width = frPrimary.Width
  
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
'On Error GoTo FormErr
    ShowStatus False
    If UnloadForm(ADOprimaryrs) = 0 Then
        'Call RedoNumbers
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
  ShowStatus True
  If ADOprimaryrs![INV PRO Posted YN] = True Then
    
    picPosted.Visible = True
    picMajor.Enabled = False
    cmdPost.Enabled = False
    grdDataGrid.Enabled = False
  Else
    If mbAddNewFlag = False Then
      CalcTotals
    End If
    picPosted.Visible = False
    picMajor.Enabled = True
    cmdPost.Enabled = True
    grdDataGrid.Enabled = True
  End If
  If IsNull(ADOprimaryrs![INV PRO Notes]) Then
    ADOprimaryrs![INV PRO Notes] = "Please write a note for this transaction"
    If mbAddNewFlag = False Then ADOprimaryrs.Update
  End If
  lblStatus.Caption = "Record: " & CStr(ADOprimaryrs.AbsolutePosition) & " of " & CStr(ADOprimaryrs.RecordCount)
  GetTextColor Me
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

Private Sub cmdAdd_Click()
'On Error GoTo AddErr  '7541954
ShowStatus True
If cmdAdd.Caption = "&Save" Then
     If Not CheckEmpty Then
        ShowStatus False
        Exit Sub
     End If
     With ADOprimaryrs
         mbAddNewFlag = False
         '.UpdateBatch adAffectAll
         '.MovePrevious
         'grdDataGrid.HoldFields
         'grdDataGrid.ReBind
         'grdDataGrid.Refresh
         cmdUpdate_Click
         ClearDatasource
         RSstatement = "SHAPE {select [INV PRO Date],[INV PRO Document No],[INV PRO Ext Document No],[INV PRO Memo],[INV PRO Notes],[INV PRO Posted YN],[INV PRO Type] from [INV Production] ORDER BY [INV PRO Document No]} AS ParentCMD APPEND ({select [INV PROD ID],[INV PROD Document No],[INV PROD Item ID],[INV PROD Description],[INV PROD Quantity],[INV PROD Cost],[INV PROD Ext Cost],[INV PROD Project] from [INV Production Detail] Order by [INV PROD Item ID] } AS ChildCMD RELATE [INV PRO Document No] TO [INV PROD Document No]) AS ChildCMD"
         OpenDB RSstatement, True
         
         NewLoad = False
         'cmdRefresh_Click
         '.MoveLast
     End With
     cmdAdd.Caption = "&Add"
     SetButtons True
     cmdPost.Enabled = True
Else
     mbAddNewFlag = True
     cmdPost.Enabled = False
     cbfields(6).Enabled = True
  With ADOprimaryrs
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
     NewLoad = True
     .AddNew
     txtfields(2) = "TempINVPRO" & AppLoginName
     txtfields(0) = FormatDate(Now)
     cbfields(6).Text = cbfields(6).List(1)
     lblStatus.Caption = "Add record"
     txtfields(3) = "-Created by " & AppLoginName
     txtfields(8) = "$0.00"
     txtfields(9) = 0
     SetButtons False
  End With
'  cmdAdd.Caption = "&Save"
End If
  ShowStatus False
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
'     If Not DataDelete(ADOprimaryrs, Me, True) Then
'     End If
Dim DocNo As String 'picStatBox

DocNo = txtfields(2).Text

        ShowStatus True
        ClearDatasource
        db.Execute "DELETE FROM [INV Production] WHERE [" & txtfields(2).DataField & "]='" & DocNo & "'"
            MsgBox lblTop & "[" & DocNo & "] has been deleted.", vbInformation, "Information"
        RSstatement = "SHAPE {select [INV PRO Date],[INV PRO Document No],[INV PRO Ext Document No],[INV PRO Memo],[INV PRO Notes],[INV PRO Posted YN],[INV PRO Type] from [INV Production] ORDER BY [INV PRO Document No]} AS ParentCMD APPEND ({select [INV PROD ID],[INV PROD Document No],[INV PROD Item ID],[INV PROD Description],[INV PROD Quantity],[INV PROD Cost],[INV PROD Ext Cost],[INV PROD Project] from [INV Production Detail] Order by [INV PROD Item ID] } AS ChildCMD RELATE [INV PRO Document No] TO [INV PROD Document No]) AS ChildCMD"
        OpenDB RSstatement
        
        ShowStatus False
End Sub

Private Sub cmdRefresh_Click()
    RefreshButton ADOprimaryrs, grdDataGrid
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
Case 0
    INV_ITEM
Case 2   'Get the type of account for the selected row
    COA_grdDataGrid_Butt
Case 7   'Select the project that have been working on
    Proj_Projects
End Select

If grdOnAddNew = True And grdDataGrid.Columns(0) <> "" Then NewgrdDatagrid
grdDataGrid_AfterColEdit 0
Exit Sub
Error_ButtClick:
    MsgBox "Please click the Table box before clicking the button"
End Sub
Private Sub grdDataGrid_AfterColEdit(ByVal ColIndex As Integer)
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
Dim TempTotal As Integer
Dim TempCurr As Currency
  Dim CalcTable  As ADODB.Recordset
  Set CalcTable = New ADODB.Recordset
  TempTotal = 0
  TempCurr = 0
  CalcTable.Open "SELECT [INV PROD Quantity],[INV PROD Cost] FROM [INV Production Detail] WHERE [INV PROD Document No]=" & ADOprimaryrs![INV PRO Document No], db, adOpenStatic, adLockOptimistic, adCmdText
  With CalcTable
    cbfields(6).Enabled = True
    If .RecordCount = 0 Then
      txtfields(8) = "$0.00"
      txtfields(9) = 0
      Exit Sub
    End If
    cbfields(6).Enabled = False
    .MoveFirst
    
    Do While Not .EOF
       'MsgBox ![INV PROD Quantity]
       TempTotal = TempTotal + ![INV PROD Quantity]
       TempCurr = TempCurr + (![INV PROD Quantity] * ![INV PROD Cost])
       .MoveNext
    Loop
  .Close
  End With
  Set CalcTable = Nothing
  txtfields(8) = FormatCurr(TempCurr)
  txtfields(9) = TempTotal
End Sub

Private Sub NewgrdDatagrid()
    NewLoad = True
     NewRowForDataGrid ADOprimaryrs, grdDataGrid, "INV PRO Notes", txtfields(3).Text
    grdOnAddNew = False
    NewLoad = False
'     With ADOprimaryrs
'        ADOprimaryrs("INV PRO Notes") = txtFields(3) & ""
'        txtFields_LostFocus 3
'        txtFields_LostFocus 5
'        ADOprimaryrs.Update
'        If Not (.BOF Or .EOF) Then
'           mvBookMark = .Bookmark
'        End If
'        Set grdDataGrid.DataSource = Nothing
'            ADOprimaryrs.UpdateBatch adAffectAll
'            ADOprimaryrs.Requery
'        Set grdDataGrid.DataSource = ADOprimaryrs("ChildCMD").UnderlyingValue
'        If mvBookMark > 0 Then
'           ADOprimaryrs.Bookmark = mvBookMark
'        Else
'           ADOprimaryrs.MoveFirst
'        End If
'    End With
'    grdOnAddNew = False
End Sub

Private Sub grdDataGrid_Error(ByVal DataError As Integer, Response As Integer)
    If DataGridKnownError(DataError) Then
        Response = 0
    End If
End Sub

Private Sub grdDataGrid_GotFocus()
Dim CreateOrder As Integer
    If mbAddNewFlag = True Then
        cmdAdd.SetFocus
        CreateOrder = MsgBox("This Request will save the data to the database? Are sure to continue", vbYesNo, "Save Quote")
        If CreateOrder = vbNo Then Exit Sub
        cmdAdd_Click
    End If
End Sub

Private Sub grdDataGrid_LostFocus()
    SendKeys ("{LEFT}")
    If txtfields(3).Enabled = True Then txtfields(3).SetFocus
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


'Private Sub txtFields_Change(Index As Integer)
'If mbAddNewFlag = True Then Exit Sub
'Select Case Index
'Case 2
    'MsgBox txtFields(1) & "   " & txtFields(2)
'    If txtFields(2) = "TempINVPRO" & AppLoginName Then
'        ADOprimaryrs![INV PRO Ext Document No] = AppLoginName & Format(Now, "MMdd") & Format(txtFields(1), "000")
'        txtFields(2) = AppLoginName & Format(Now, "MMdd") & Format(txtFields(1), "000")
'        ADOprimaryrs.Update
'    End If
'End Select
'End Sub

Private Sub txtFields_LostFocus(Index As Integer)
If NewLoad = True Then Exit Sub
Select Case Index
Case 0
    txtfields(0) = FormatDate(Now)
Case 3
    If InStr(1, txtfields(Index), "-Created by " & AppLoginName, vbTextCompare) Then Exit Sub
    txtfields(Index) = txtfields(Index) & "-Created by " & AppLoginName
Case 5
    If Len(Trim(txtfields(Index))) = 0 Then
       MsgBox "Please enter the reason for the Production", vbCritical, "Error"
       txtfields(5).SetFocus
    End If
End Select

End Sub


