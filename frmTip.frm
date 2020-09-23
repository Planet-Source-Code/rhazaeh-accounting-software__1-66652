VERSION 5.00
Begin VB.Form frmTip 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Message of the Day"
   ClientHeight    =   3300
   ClientLeft      =   5430
   ClientTop       =   4080
   ClientWidth     =   4590
   Icon            =   "frmTip.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdNextTip 
      Caption         =   "&Next Tip"
      Height          =   375
      Left            =   1980
      TabIndex        =   2
      Top             =   2880
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   2715
      Left            =   80
      Picture         =   "frmTip.frx":030A
      ScaleHeight     =   2655
      ScaleWidth      =   4395
      TabIndex        =   1
      Top             =   70
      Width           =   4455
      Begin VB.Label lblBy 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "by"
         DataField       =   "By"
         Height          =   255
         Left            =   180
         TabIndex        =   5
         Top             =   2400
         Width           =   4095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "I want you to know that,"
         Height          =   255
         Left            =   540
         TabIndex        =   4
         Top             =   180
         Width           =   2655
      End
      Begin VB.Label lblTipText 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Message"
         DataField       =   "Text"
         Height          =   1755
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   4215
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3300
      TabIndex        =   0
      Top             =   2880
      Width           =   1215
   End
End
Attribute VB_Name = "frmTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents AdoRs As ADODB.Recordset
Attribute AdoRs.VB_VarHelpID = -1
Dim db As ADODB.Connection

Dim LoadType As Integer

Private Sub AdoRs_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
If AdoRs.RecordCount = 0 Then Exit Sub
If AdoRs.EOF Then AdoRs.MoveFirst
  lblBy.Caption = "Record: " & CStr(AdoRs.AbsolutePosition) & "/" & CStr(AdoRs.RecordCount) & "   By: " & AdoRs!By
End Sub

Private Sub cmdNextTip_Click()
'On Error Resume Next
If AdoRs.RecordCount > 0 Then
    AdoRs.MoveNext
End If
End Sub

Private Sub cmdOK_Click()
    fMainForm.Timer1.Enabled = False
    Unload Me
End Sub

Public Sub TypeOfWise(WiseType As Integer)

    ' Seed Rnd
    Randomize

    LoadType = WiseType
Select Case WiseType
Case 0
    'Funny Message about everything
    cmdNextTip.Visible = False
    cmdNextTip.Caption = "&Next Message"
    Me.Caption = "Message of the day"
    Label1.Caption = "I want you to know that,"
Case 1
    'Bulletin Board
    cmdNextTip.Visible = True
    cmdNextTip.Caption = "&Next Bulletin"
    Me.Caption = "Bulletin Board"
    Label1.Caption = "I want you to know that,"
Case 2
    'accounting tips
    cmdNextTip.Visible = True
    cmdNextTip.Caption = "&Next Tip"
    Me.Caption = "Tips of the day"
    Label1.Caption = "I want you to know that,"
End Select
    frmTip.Show vbModal
End Sub

Private Sub Form_Load()
On Error GoTo FormErr
  Set db = New ADODB.Connection
  db.CursorLocation = adUseClient
  db.Open gblBasicADOProvider

  Set AdoRs = New ADODB.Recordset
  Select Case LoadType
  Case 0, 1, 2
    AdoRs.Open "SELECT * FROM [Wise Advice] WHERE Type=" & LoadType, db, adOpenKeyset, adLockReadOnly, adCmdText
  Case Else
    AdoRs.Open "SELECT * FROM [Wise Advice]", db, adOpenKeyset, adLockReadOnly, adCmdText
  End Select
  
  Set lblTipText.DataSource = AdoRs
  
    ' Seed Rnd
    Randomize
    Dim rsLocation As Integer
    If AdoRs.RecordCount > 0 Then
        'AdoRs.MoveFirst
        'do it randomly
        rsLocation = (Int((AdoRs.RecordCount * Rnd) + 1))
        AdoRs.Move (rsLocation)
    End If
Exit Sub
FormErr:
  MsgBox Err.Description
  ShowStatus False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set lblTipText.DataSource = Nothing
    AdoRs.Close
    Set AdoRs = Nothing
End Sub
