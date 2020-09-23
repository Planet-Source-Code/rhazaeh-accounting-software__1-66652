VERSION 5.00
Begin VB.Form Main_Login 
   BackColor       =   &H8000000B&
   BorderStyle     =   0  'None
   ClientHeight    =   2595
   ClientLeft      =   4305
   ClientTop       =   3750
   ClientWidth     =   7380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   Tag             =   "1055"
   Begin VB.Frame frlogin 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   40
      Visible         =   0   'False
      Width           =   4815
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1680
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1080
         Width           =   2325
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   360
         Left            =   720
         TabIndex        =   3
         Tag             =   "1058"
         Top             =   1800
         Width           =   1140
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   360
         Left            =   2880
         TabIndex        =   4
         Tag             =   "1059"
         Top             =   1800
         Width           =   1140
      End
      Begin VB.ComboBox dtUserName 
         Height          =   315
         Left            =   1680
         Style           =   1  'Simple Combo
         TabIndex        =   1
         Text            =   "dtUserName"
         Top             =   480
         Width           =   2325
      End
      Begin VB.CommandButton cmdUpdatedua 
         Height          =   280
         Index           =   2
         Left            =   3600
         Picture         =   "Main_Login.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Update the Ship Via"
         Top             =   480
         Width           =   375
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "Main_Login.frx":04B6
         Top             =   1080
         Width           =   480
      End
      Begin VB.Label lblLabels 
         Caption         =   "&User Name:"
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   6
         Tag             =   "1056"
         Top             =   480
         Width           =   840
      End
      Begin VB.Label lblLabels 
         Caption         =   "&Password:"
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   5
         Tag             =   "1057"
         Top             =   1080
         Width           =   840
      End
   End
   Begin VB.Frame fraMainFrame 
      BackColor       =   &H8000000A&
      Height          =   2415
      Left            =   120
      TabIndex        =   7
      Top             =   40
      Width           =   7140
      Begin VB.PictureBox picLogo 
         Height          =   2220
         Left            =   3720
         Picture         =   "Main_Login.frx":07C0
         ScaleHeight     =   2160
         ScaleWidth      =   3300
         TabIndex        =   15
         Top             =   140
         Width           =   3360
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   1320
         Top             =   1320
      End
      Begin VB.Label Label1 
         Caption         =   "NoSecret Accounting"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label Label1 
         Caption         =   "(C)opyright 2000, NoSecret Accounting Tech."
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   13
         Top             =   1800
         Width           =   3255
      End
      Begin VB.Label Label1 
         Caption         =   "All Rights Reserved"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   12
         Top             =   2040
         Width           =   3255
      End
      Begin VB.Label Label1 
         Caption         =   "Version v1.0 Build 1.8.19 Beta"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "ADO 2.1 Data Access Methods"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "MDB Back End Database"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Image Image2 
         Height          =   630
         Left            =   2640
         Picture         =   "Main_Login.frx":1DD0
         Top             =   1080
         Width           =   660
      End
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   6  'Inside Solid
      Height          =   2595
      Left            =   0
      Top             =   0
      Width           =   7380
   End
End
Attribute VB_Name = "Main_Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpbuffer As String, nSize As Long) As Long

Public OK As Boolean
Dim DoneLoad As Boolean

Private Sub cmdUpdatedua_Click(Index As Integer)
ShowStatus True
    ComboInit dtUserName, lblLabels(0), "SELECT [EMP Name] FROM [EMP Employees]"
ShowStatus False
End Sub

Private Sub SecondLoad()
    Me.Caption = "No Secret Accounting VB Login"
    Me.Enabled = True
    Shape1.Width = 5050
    Me.Width = Shape1.Width
    fraMainFrame.Visible = False
    frlogin.Visible = True
    frlogin.ZOrder 0

End Sub

Private Sub FirstLoad()

    'load database
    frlogin.Left = 120
    frlogin.Top = 40
    fraMainFrame.Left = frlogin.Left
    fraMainFrame.Top = frlogin.Top
    Me.Height = 2600
    Shape1.Width = 7380
    Me.Width = Shape1.Width
    fraMainFrame.Visible = True
    fraMainFrame.ZOrder 0
    Me.Enabled = False
    Me.Refresh
    GetCompany

    'prepare for the login form
    Dim sBuffer As String
    Dim lSize As Long
    ComboInit dtUserName, lblLabels(0), "SELECT [EMP ID] FROM [EMP Employees]"

    sBuffer = Space$(255)
    lSize = Len(sBuffer)
    Call GetUserName(sBuffer, lSize)
    If lSize > 0 Then
        dtUserName.Text = Left$(sBuffer, lSize)
    Else
        dtUserName.Text = vbNullString
    End If
    
    DoneLoad = True
End Sub
Private Sub cmdCancel_Click()
    OK = False
    Me.Hide
End Sub


Private Sub cmdOK_Click()
    'Start the Application, First load the Main and Splash Screens
    'Set fMainForm = New Main_Menu
    'Load fMainForm
    'ToDo: create test for correct password
    'check for correct password
    If Trim(dtUserName.Text) <> "" Then
        OK = True
        AppLoginName = dtUserName.Text
        Me.Hide
    Else
        MsgBox "Invalid Password, try again!", , "Login"
        txtPassword.SetFocus
        txtPassword.SelStart = 0
        txtPassword.SelLength = Len(txtPassword.Text)
    End If
        fMainForm.txtLogon = "Logon As " & AppLoginName
End Sub

Private Sub Form_Load()
    DoneLoad = False
    Timer1.Enabled = True
    Me.Left = (Screen.Width - Me.ScaleWidth) / 2
    Me.Top = (Screen.Height - Me.ScaleHeight) / 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Timer1.Enabled = False
    Set Main_Login = Nothing
End Sub

Private Sub Timer1_Timer()
If DoneLoad = False Then
    FirstLoad
Else
    SecondLoad
    Me.Refresh
    Timer1.Enabled = False
    cmdOK.SetFocus
End If
End Sub
