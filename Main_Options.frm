VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Begin VB.Form Main_Options 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Properties"
   ClientHeight    =   6285
   ClientLeft      =   4350
   ClientTop       =   3270
   ClientWidth     =   9405
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   9405
   ShowInTaskbar   =   0   'False
   Tag             =   "1060"
   Begin VB.TextBox txtLocked 
      Appearance      =   0  'Flat
      DataField       =   "BackColor"
      Height          =   285
      Index           =   0
      Left            =   10920
      TabIndex        =   18
      Top             =   2520
      Width           =   1200
   End
   Begin VB.TextBox txtLocked 
      Appearance      =   0  'Flat
      DataField       =   "ForeColor"
      Height          =   285
      Index           =   2
      Left            =   10920
      TabIndex        =   17
      Top             =   4440
      Width           =   1200
   End
   Begin VB.TextBox txtLocked 
      DataField       =   "Font"
      Height          =   285
      Index           =   3
      Left            =   10920
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   3960
      Width           =   1200
   End
   Begin VB.TextBox txtLocked 
      DataField       =   "BorderStyle"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   10920
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   3480
      Width           =   1200
   End
   Begin VB.TextBox txtLocked 
      DataField       =   "Appearance"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   10920
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   3000
      Width           =   1200
   End
   Begin VB.TextBox txtLocked 
      DataField       =   "Appearance"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   10920
      TabIndex        =   13
      Top             =   2040
      Width           =   1215
   End
   Begin ComctlLib.TreeView tvwDB 
      Height          =   5655
      Left            =   0
      TabIndex        =   12
      Top             =   600
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   9975
      _Version        =   327682
      Style           =   7
      Appearance      =   1
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   5640
      TabIndex        =   0
      Tag             =   "1067"
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6840
      TabIndex        =   1
      Tag             =   "1066"
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8040
      TabIndex        =   2
      Tag             =   "1065"
      Top             =   5880
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog cmnDialog 
      Left            =   6720
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frOptions 
      Height          =   5295
      Index           =   1
      Left            =   3360
      TabIndex        =   19
      Tag             =   "1061"
      Top             =   480
      Width           =   6000
      Begin VB.Frame Frame1 
         Caption         =   "Sample"
         Height          =   2775
         Left            =   120
         TabIndex        =   36
         Top             =   2400
         Width           =   5775
         Begin VB.PictureBox Picture2 
            Height          =   2415
            Left            =   120
            ScaleHeight     =   2355
            ScaleWidth      =   5475
            TabIndex        =   37
            Top             =   240
            Width           =   5535
            Begin VB.Frame Frame3 
               Height          =   2295
               Left            =   120
               TabIndex        =   41
               Top             =   0
               Width           =   2655
               Begin VB.PictureBox Picture1 
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000A&
                  BorderStyle     =   0  'None
                  Enabled         =   0   'False
                  ForeColor       =   &H80000008&
                  Height          =   615
                  Left            =   960
                  ScaleHeight     =   615
                  ScaleWidth      =   1455
                  TabIndex        =   44
                  Top             =   1560
                  Width           =   1455
                  Begin VB.TextBox txtSample 
                     Height          =   285
                     Index           =   1
                     Left            =   0
                     TabIndex        =   45
                     Text            =   "Disabled Input Box"
                     Top             =   120
                     Width           =   1455
                  End
               End
               Begin VB.TextBox txtSample 
                  Height          =   285
                  Index           =   2
                  Left            =   960
                  TabIndex        =   43
                  Text            =   "Standard Input Box"
                  Top             =   360
                  Width           =   1455
               End
               Begin VB.TextBox txtSample 
                  Height          =   285
                  Index           =   0
                  Left            =   960
                  Locked          =   -1  'True
                  TabIndex        =   42
                  Text            =   "Locked Input Box"
                  Top             =   1080
                  Width           =   1455
               End
               Begin VB.Label lblLabels 
                  Caption         =   "Side Label"
                  Height          =   285
                  Index           =   1
                  Left            =   120
                  TabIndex        =   47
                  Top             =   360
                  Width           =   855
               End
               Begin VB.Label lblfields 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Above Label"
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Index           =   0
                  Left            =   960
                  TabIndex        =   46
                  Top             =   840
                  Width           =   1455
               End
            End
            Begin VB.Frame Frame2 
               Caption         =   "Frame Appearance"
               Height          =   975
               Left            =   2880
               TabIndex        =   40
               Top             =   1320
               Width           =   2535
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Check Box Apperance"
               Height          =   255
               Left            =   3480
               TabIndex        =   39
               Top             =   840
               Width           =   1935
            End
            Begin VB.OptionButton Option1 
               Alignment       =   1  'Right Justify
               Caption         =   "Option Appearance"
               Height          =   255
               Left            =   3600
               TabIndex        =   38
               Top             =   360
               Width           =   1815
            End
         End
      End
      Begin VB.PictureBox picLocked 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   1320
         ScaleHeight     =   255
         ScaleWidth      =   1170
         TabIndex        =   26
         Top             =   1080
         Width           =   1200
      End
      Begin VB.CommandButton cmdLoked 
         Height          =   285
         Index           =   0
         Left            =   2520
         TabIndex        =   25
         Top             =   1080
         Width           =   255
      End
      Begin VB.CommandButton cmdLoked 
         Height          =   285
         Index           =   1
         Left            =   5520
         TabIndex        =   24
         Top             =   1080
         Width           =   255
      End
      Begin VB.CommandButton cmdLoked 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   2520
         TabIndex        =   23
         Top             =   1800
         Width           =   255
      End
      Begin VB.PictureBox picLocked 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   1320
         ScaleHeight     =   255
         ScaleWidth      =   1155
         TabIndex        =   22
         Top             =   1800
         Width           =   1185
      End
      Begin VB.CommandButton cmdLoked 
         Height          =   285
         Index           =   3
         Left            =   5520
         TabIndex        =   21
         Top             =   1440
         Width           =   255
      End
      Begin VB.CommandButton cmdLoked 
         Height          =   285
         Index           =   4
         Left            =   2520
         TabIndex        =   20
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label lblLocked 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "BackColor"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   35
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblLocked 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Appearance"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   3240
         TabIndex        =   34
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblLocked 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ForeColor"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   33
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblLocked 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Font"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   3240
         TabIndex        =   32
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label lblLocked 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "BorderStyle"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   31
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label lblTops 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Locked"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   360
         Width           =   5775
      End
      Begin VB.Label lbl 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   1
         Left            =   4320
         TabIndex        =   29
         Top             =   1080
         Width           =   1200
      End
      Begin VB.Label lbl 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   4
         Left            =   1320
         TabIndex        =   28
         Top             =   1440
         Width           =   1200
      End
      Begin VB.Label lbl 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   3
         Left            =   4320
         TabIndex        =   27
         Top             =   1440
         Width           =   1200
      End
   End
   Begin VB.Frame frProperties 
      Height          =   5295
      Left            =   3360
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   6000
      Begin VB.CommandButton Command3 
         Caption         =   "None"
         Height          =   285
         Left            =   4920
         TabIndex        =   10
         Top             =   1560
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Default"
         Height          =   285
         Left            =   4200
         TabIndex        =   9
         Top             =   1560
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Height          =   285
         Left            =   5280
         Picture         =   "Main_Options.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox txtLocked 
         DataField       =   "BackColor"
         Height          =   285
         Index           =   6
         Left            =   1680
         TabIndex        =   7
         Top             =   1200
         Width           =   3615
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Main_Options.frx":04B6
         Left            =   1680
         List            =   "Main_Options.frx":04C0
         TabIndex        =   6
         Text            =   "Big"
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "The change will take effect after you restart the application."
         Height          =   495
         Left            =   1680
         TabIndex        =   11
         Top             =   1920
         Width           =   3615
      End
      Begin VB.Label lblStartup 
         Caption         =   "ToolBar Size"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   5
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblStartup 
         Caption         =   "ToolBar Picture"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   4
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Image imgOptions 
         Height          =   1605
         Index           =   0
         Left            =   1440
         Picture         =   "Main_Options.frx":04D0
         Top             =   2520
         Visible         =   0   'False
         Width           =   3825
      End
      Begin VB.Image imgOptions 
         Height          =   1500
         Index           =   1
         Left            =   1440
         Picture         =   "Main_Options.frx":14612
         Top             =   2520
         Visible         =   0   'False
         Width           =   3780
      End
   End
End
Attribute VB_Name = "Main_Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TemptConnection As String
Dim TempgblADOProvider As String
Dim dbTemp As ADODB.Connection
Dim ADOprimaryrs As ADODB.Recordset
Dim mNode As Node


Private Sub cmdLoked_Click(Index As Integer)
On Error GoTo FormErr
Select Case Index
Case 0
    cmnDialog.ShowColor
    txtLocked(Index).Text = cmnDialog.Color
    txtLocked(Index).SetFocus
    picLocked(Index).BackColor = cmnDialog.Color
    cmdApply.Enabled = True
Case 1
    If txtLocked(Index).Text = 1 Then
        txtLocked(Index).Text = 0
    Else
        txtLocked(Index).Text = 1
    End If
    txtLocked(Index).SetFocus
    txtLocked(Index).Appearance = txtLocked(Index).Text
    cmdApply.Enabled = True
Case 2
    cmnDialog.ShowColor
    txtLocked(Index).Text = cmnDialog.Color
    txtLocked(Index).SetFocus
    picLocked(Index).BackColor = cmnDialog.Color
    cmdApply.Enabled = True
Case 3
    cmnDialog.ShowFont
    txtLocked(Index).Text = cmnDialog.FontName
    txtLocked(Index).SetFocus
    picLocked(Index).BackColor = cmnDialog.Color
    cmdApply.Enabled = True
Case 4
    If txtLocked(Index).Text = 1 Then
        txtLocked(Index).Text = 0
    Else
        txtLocked(Index).Text = 1
    End If
    txtLocked(Index).SetFocus
    txtLocked(Index).BorderStyle = txtLocked(Index).Text
    cmdApply.Enabled = True
End Select

CheckColor

Exit Sub
FormErr:
  'MsgBox Err.Description
  ShowStatus False
End Sub


Private Sub Combo1_Click()
Dim sFile As String

If Combo1.Text = "Big" Then
    sFile = "1"
    imgOptions(0).Visible = True
    imgOptions(1).Visible = False
Else
    sFile = "0"
    imgOptions(0).Visible = False
    imgOptions(1).Visible = True
End If
    txtLocked(5).Text = sFile
    txtLocked(5).SetFocus
    cmdApply.Enabled = True
End Sub

Private Sub Command1_Click()
    Dim sFile As String
    
    sFile = ""
    With fMainForm.dlgCommonDialog
        .DialogTitle = "Load Picture For ToolBar"
        .CancelError = False
        .Flags = cdlOFNPathMustExist + cdlOFNOverwritePrompt
        .Filter = "All Files (*.BMP;*.Gif;*.JPG)|*.BMP;*.Gif;*.JPG"
        .ShowOpen
        If Len(.FileName) = 0 Then
            ShowStatus False
            Exit Sub
        End If
        sFile = .FileName
    End With
    If sFile <> "" Then
        txtLocked(6).Text = sFile
        txtLocked(6).SetFocus
        'Set fMainForm.CoolBar1.Picture = LoadPicture(sFile)
        'fMainForm.CoolBar1.Refresh
        'fMainForm.tbToolBar1.RestoreToolbar
        cmdApply.Enabled = True
    End If
End Sub

Private Sub Command2_Click()
        txtLocked(6).Text = "Default"
        txtLocked(6).SetFocus
        cmdApply.Enabled = True
End Sub

Private Sub Command3_Click()
        txtLocked(6).Text = "None"
        txtLocked(6).SetFocus
        cmdApply.Enabled = True
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim Relatif As String

'On Error GoTo FormErr

Dim SQLstatement As String
     
      Set dbTemp = New ADODB.Connection
      dbTemp.CursorLocation = adUseClient
      dbTemp.Open gblBasicADOProvider
      
      SQLstatement = "SELECT * FROM Properties"
      Set ADOprimaryrs = New ADODB.Recordset
      ADOprimaryrs.Open SQLstatement, dbTemp, adOpenKeyset, adLockOptimistic, adCmdText
      ADOprimaryrs.MoveFirst
    
    Set mNode = tvwDB.Nodes.Add()
        Relatif = "NoSecret Properties"
        mNode.Text = Relatif
        mNode.Tag = Relatif
        mNode.Key = Relatif
    
    Do While Not ADOprimaryrs.EOF
        Set mNode = tvwDB.Nodes.Add(Relatif, tvwChild)
           mNode.Text = ADOprimaryrs![Type]
           mNode.Tag = ADOprimaryrs![Type]
           mNode.Key = ADOprimaryrs![Type]
        ADOprimaryrs.MoveNext
    Loop
    ADOprimaryrs.MoveFirst
    
        Dim Ctrl As Control
        For Each Ctrl In Me.Controls
          If TypeOf Ctrl Is TextBox Then
             Set Ctrl.DataSource = ADOprimaryrs
           End If
        Next
        CheckColor

    'tbsOptions.SelectedItem = tbsOptions.Tabs(1)
    'picOptions.Left = 240
    'picOptions.Top = 960
    'picOptions.ZOrder 0
    'picOptions.Enabled = True
    
    'Me.Width = 6330
    'Me.Height = 5895
    Sample
    tvwDB.Nodes("" & Relatif & "").Expanded = True
Exit Sub
FormErr:
    MsgBox "Missing Database. Please contact the supplier", vbCritical, "Error"
End Sub

Private Sub CheckColor()
        picLocked(0).BackColor = txtLocked(0).Text
        picLocked(2).BackColor = txtLocked(2).Text
        
        Dim Ctrl As Control
             
             'txtSample.BackColor = txtLocked(0).Text
             'txtSample.Appearance = txtLocked(1).Text
                If txtLocked(1).Text = 1 Then
                    lbl(1).Caption = "3D"
                Else
                    lbl(1).Caption = "Flat"
                End If
             'txtSample.BorderStyle = txtLocked(4).Text
                If txtLocked(4).Text = 1 Then
                    lbl(4).Caption = "Fixed Single"
                Else
                    lbl(4).Caption = "None"
                End If
             'txtSample.Font = txtLocked(3).Text
             lbl(3).Caption = txtLocked(3).Text
             'txtSample.ForeColor = txtLocked(2).Text
End Sub
Private Sub cmdApply_Click()
    ADOprimaryrs.Update
    cmdApply.Enabled = False
    Sample
End Sub

Private Sub cmdCancel_Click()
    ADOprimaryrs.CancelUpdate
    Unload Me
End Sub

Private Sub cmdOK_Click()
    ADOprimaryrs.UpdateBatch adAffectAll
    Unload Me
End Sub

Private Sub Sample()
Dim ADOsecondary As ADODB.Recordset

    Set ADOsecondary = New ADODB.Recordset
    ADOsecondary.Open "SELECT * FROM Properties", dbTemp, adOpenKeyset, adLockReadOnly, adCmdText
  
  Dim Ctrltxt As TextBox
With ADOsecondary
  For Each Ctrltxt In Me.txtSample
    If Ctrltxt.Enabled = False Then
        .MoveFirst
        .Find "Type='Disabled Input Box'"
    ElseIf Ctrltxt.Locked = True Then
        .MoveFirst
        .Find "Type='Locked Input Box'"
    Else
        .MoveFirst
        .Find "Type='Standard Input Box'"
    End If
    
        Ctrltxt.BackColor = ADOsecondary("BackColor").Value
        Ctrltxt.Appearance = ADOsecondary("Appearance").Value
        Ctrltxt.BorderStyle = ADOsecondary("BorderStyle").Value
        Ctrltxt.Font = ADOsecondary("Font").Value
        Ctrltxt.ForeColor = ADOsecondary("ForeColor").Value
 Next
 
   'For labels
  Dim Ctrllbl As Control
  For Each Ctrllbl In Me.Controls
     If Ctrllbl.Name = "lblfields" Then
         .MoveFirst
         .Find "Type='Above Label'"
         Ctrllbl.BackStyle = 1
     ElseIf Ctrllbl.Name = "lblLabels" Then
         .MoveFirst
         .Find "Type='Side Label'"
         Ctrllbl.BackStyle = 1
     Else
         GoTo JumpLoop
     End If
         Ctrllbl.Appearance = ADOsecondary("Appearance").Value
         Ctrllbl.BorderStyle = ADOsecondary("BorderStyle").Value
         Ctrllbl.Font = ADOsecondary("Font").Value
         Ctrllbl.ForeColor = ADOsecondary("ForeColor").Value
         Ctrllbl.BackColor = ADOsecondary("BackColor").Value
JumpLoop:
  Next
         .MoveFirst
         .Find "Type='Interface'"
         
  Dim CtrlInter As Control
  For Each CtrlInter In Me.Controls
    If TypeOf CtrlInter Is Frame And CtrlInter.Name <> "frProperties" And CtrlInter.Name <> "frOptions" And CtrlInter.Name <> "Frame1" Then
        CtrlInter.Appearance = ADOsecondary("Appearance").Value
        CtrlInter.BorderStyle = ADOsecondary("BorderStyle").Value
        CtrlInter.Font = ADOsecondary("Font").Value
        CtrlInter.ForeColor = ADOsecondary("ForeColor").Value
        CtrlInter.BackColor = ADOsecondary("BackColor").Value
    ElseIf TypeOf CtrlInter Is PictureBox And CtrlInter.Name <> "picLocked" Then
        'Picture2.Appearance = ADOsecondary("Appearance").Value
        'Picture2.BorderStyle = ADOsecondary("BorderStyle").Value
        CtrlInter.Font = ADOsecondary("Font").Value
        CtrlInter.ForeColor = ADOsecondary("ForeColor").Value
        CtrlInter.BackColor = ADOsecondary("BackColor").Value
    ElseIf TypeOf CtrlInter Is CheckBox Or TypeOf CtrlInter Is OptionButton Then
        CtrlInter.Appearance = ADOsecondary("Appearance").Value
        'Option1.BorderStyle = ADOsecondary("BorderStyle").Value
        CtrlInter.Font = ADOsecondary("Font").Value
        CtrlInter.ForeColor = ADOsecondary("ForeColor").Value
        CtrlInter.BackColor = ADOsecondary("BackColor").Value
    End If
  Next
End With
    ADOsecondary.Close
    Set ADOsecondary = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ADOprimaryrs.CancelUpdate
    ADOprimaryrs.Close
    Set ADOprimaryrs = Nothing
    dbTemp.Close
    Set dbTemp = Nothing
    Set Main_Options = Nothing
End Sub

Private Sub tvwDB_NodeClick(ByVal Node As ComctlLib.Node)
Select Case tvwDB.SelectedItem
Case "Startup"
    ADOprimaryrs.MoveFirst
    ADOprimaryrs.Find "[Type]='Startup'"
    'MsgBox ADOprimaryrs.EOF
    frProperties.Visible = True
    frProperties.ZOrder 0
    'tbsOptions.Visible = False
    'picOptions.Visible = False
    If txtLocked(5).Text = 1 Then
        Combo1.Text = "Big"
        imgOptions(0).Visible = True
        imgOptions(1).Visible = False
    Else
    imgOptions(0).Visible = False
    imgOptions(1).Visible = True
        Combo1.Text = "Small"
    End If
Case Else
    ADOprimaryrs.MoveFirst
    'tbsOptions.Visible = True
    'picOptions.Visible = True
    frProperties.Visible = False
    frOptions(1).ZOrder 0
  ADOprimaryrs.Find "[Type]='" & tvwDB.SelectedItem & "'"
  If Not ADOprimaryrs.EOF Then
    lblTops.Caption = tvwDB.SelectedItem
    CheckColor
  End If
End Select

End Sub
