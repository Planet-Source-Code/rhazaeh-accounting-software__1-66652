VERSION 5.00
Begin VB.Form frm_Help 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Help"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   11565
   Begin VB.Frame frHelp 
      Height          =   7215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11535
      Begin VB.TextBox txtHelp 
         Height          =   6015
         Left            =   3960
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Text            =   "frm_Help.frx":0000
         Top             =   840
         Width           =   7455
      End
      Begin VB.ComboBox cbhelp 
         Height          =   6030
         Left            =   120
         Sorted          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   840
         Width           =   3735
      End
      Begin VB.CommandButton cmdCloseHelp 
         Height          =   615
         Left            =   10800
         Picture         =   "frm_Help.frx":001C
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   180
         Width           =   615
      End
      Begin VB.Label lblhelp 
         Alignment       =   2  'Center
         Caption         =   "Help Text:"
         BeginProperty Font 
            Name            =   "Vibrocentric"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   10695
      End
   End
End
Attribute VB_Name = "frm_Help"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As ADODB.Connection
Dim Helprs As ADODB.Recordset

Private Sub cbhelp_Click()

    lblhelp = cbhelp.Text
    Helprs.MoveFirst
    Helprs.Find "[Form Name]='" & cbhelp.Text & "'"
End Sub

Private Sub cmdCloseHelp_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Helprs.CancelUpdate
    Helprs.Close
    Set Helprs = Nothing
    Set frm_Help = Nothing
End Sub

Private Sub Form_Load()
  Set db = New ADODB.Connection
  db.CursorLocation = adUseClient
  'db.Provider = "MSDataShape"
  'db.Open "Data " & gblADOProvider
  db.Open gblBasicADOProvider
        
  Set Helprs = New ADODB.Recordset
  Helprs.Open "SELECT [Form Name],[Help Text] FROM [Help Text]", db, adOpenKeyset, adLockOptimistic, adCmdText
    
  Set txtHelp.DataSource = Helprs
  txtHelp.DataField = "Help Text"
  
  ComboInit cbhelp, lblhelp, "SELECT [Form Name] FROM [Help Text]", db
End Sub
