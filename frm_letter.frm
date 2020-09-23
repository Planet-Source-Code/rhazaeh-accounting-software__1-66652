VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frm_letter 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Business Letter/Planning"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10725
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5790
   ScaleWidth      =   10725
   Begin ComctlLib.TreeView tvwDB 
      Height          =   5290
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   9340
      _Version        =   327682
      HideSelection   =   0   'False
      Indentation     =   647
      LabelEdit       =   1
      Sorted          =   -1  'True
      Style           =   7
      Appearance      =   1
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   375
      Left            =   3420
      TabIndex        =   1
      Top             =   480
      Width           =   7275
   End
   Begin RichTextLib.RichTextBox RTF 
      Height          =   4890
      Left            =   3405
      TabIndex        =   0
      Top             =   900
      Width           =   7290
      _ExtentX        =   12859
      _ExtentY        =   8625
      _Version        =   393217
      ScrollBars      =   3
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frm_letter.frx":0000
   End
   Begin VB.Label lblTop 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Business Letter/Planning"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   10695
   End
End
Attribute VB_Name = "frm_letter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'private HWND_TOPMOST = -1

Private db As ADODB.Connection
Private WithEvents ADOprimaryrs As ADODB.Recordset
Attribute ADOprimaryrs.VB_VarHelpID = -1
Private mNode As Node


Private Sub cmdUpdate_Click()
    ADOprimaryrs![11] = RTF.Text
    ADOprimaryrs.Update
End Sub

Private Sub Form_Load()
ShowStatus True
Dim gblADOProvider As String
Dim Relatif As String
 ' Declare needed variables
 ' Use the SetWindowPos API Function

  'gblADOProvider = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\letter.mdb" & ";Persist Security Info=False"
    
  Set db = New ADODB.Connection
  db.CursorLocation = adUseClient
  db.Open gblBasicADOProvider
  
  Set mNode = tvwDB.Nodes.Add()
  Relatif = "TBS' Browser"
  mNode.Text = Relatif
  mNode.Tag = Relatif
  mNode.Key = Relatif
  
Set ADOprimaryrs = New ADODB.Recordset
ADOprimaryrs.Open "SELECT [ID],[4],[0] FROM [Letter] WHERE [5] is Null", db, adOpenForwardOnly, adLockReadOnly, adCmdText
    With ADOprimaryrs
        If .RecordCount > 0 Then
        .MoveFirst
            Do While Not .EOF
                Set mNode = tvwDB.Nodes.Add("TBS' Browser", tvwChild)
                Relatif = "TBS" & Str(ADOprimaryrs![4])
                mNode.Tag = ADOprimaryrs![ID]
                mNode.Key = Relatif
                mNode.Text = ADOprimaryrs![0]
            .MoveNext
            Loop
        End If
    End With
ADOprimaryrs.Close
Set ADOprimaryrs = Nothing

Set ADOprimaryrs = New ADODB.Recordset
ADOprimaryrs.Open "SELECT [ID],[4],[5],[0] FROM [Letter] WHERE [6] is Null and [5]<> null", db, adOpenForwardOnly, adLockReadOnly, adCmdText
    With ADOprimaryrs
        If .RecordCount > 0 Then
        .MoveFirst
            Do While Not .EOF
                Relatif = "TBS" & Str(ADOprimaryrs![4])
                Set mNode = tvwDB.Nodes.Add(Relatif, tvwChild)
                Relatif = Relatif & Str(ADOprimaryrs![5])
                mNode.Tag = ADOprimaryrs![ID]
                mNode.Key = Relatif
                'mNode.Text = ADOprimaryRS![id]
                mNode.Text = ADOprimaryrs![0]
            .MoveNext
            Loop
        End If
    End With
ADOprimaryrs.Close
Set ADOprimaryrs = Nothing

Set ADOprimaryrs = New ADODB.Recordset
ADOprimaryrs.Open "SELECT [ID],[4],[5],[6],[0] FROM [Letter] WHERE [7] is Null and [6]<> null", db, adOpenForwardOnly, adLockReadOnly, adCmdText
    With ADOprimaryrs
        If .RecordCount > 0 Then
        .MoveFirst
            Do While Not .EOF
                Relatif = "TBS" & Str(ADOprimaryrs![4]) & Str(ADOprimaryrs![5])
                Set mNode = tvwDB.Nodes.Add(Relatif, tvwChild)
                Relatif = Relatif & Str(ADOprimaryrs![6])
                mNode.Tag = ADOprimaryrs![ID]
                mNode.Key = Relatif
                mNode.Text = ADOprimaryrs![0]
            .MoveNext
            Loop
        End If
    End With
ADOprimaryrs.Close
Set ADOprimaryrs = Nothing

Set ADOprimaryrs = New ADODB.Recordset
ADOprimaryrs.Open "SELECT [ID],[4],[5],[6],[7],[0] FROM [Letter] WHERE [8] is Null and [7]<> null", db, adOpenForwardOnly, adLockReadOnly, adCmdText
    With ADOprimaryrs
        If .RecordCount > 0 Then
        .MoveFirst
            Do While Not .EOF
                Relatif = "TBS" & Str(ADOprimaryrs![4]) & Str(ADOprimaryrs![5]) & Str(ADOprimaryrs![6])
                Set mNode = tvwDB.Nodes.Add(Relatif, tvwChild)
                Relatif = Relatif & Str(ADOprimaryrs![7])
                mNode.Tag = ADOprimaryrs![ID]
                mNode.Key = Relatif
                mNode.Text = ADOprimaryrs![0]
            .MoveNext
            Loop
        End If
    End With
ADOprimaryrs.Close
Set ADOprimaryrs = Nothing
tvwDB.Nodes("TBS' Browser").Expanded = True
ShowStatus False

End Sub

Private Sub Form_Resize()
On Error GoTo exit_sub
    tvwDB.Height = Me.Height - 400
    RTF.Height = Me.Height - RTF.Top - 400
    RTF.Width = Me.Width - RTF.Left - 100
    cmdUpdate.Width = Me.Width - RTF.Left - 200
    lblTop.Width = Me.Width - 155 '10845-10695
exit_sub:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If ADOprimaryrs Is Nothing Then
    Else
        ADOprimaryrs.Close
        Set ADOprimaryrs = Nothing
    End If
    
    db.Close
    Set db = Nothing
End Sub

Private Sub tvwDB_Expand(ByVal Node As ComctlLib.Node)
    tvwDB.Nodes(Node.Key).Sorted = True
End Sub

Private Sub tvwDB_NodeClick(ByVal Node As ComctlLib.Node)
ShowStatus True
If tvwDB.SelectedItem.Child Is Nothing Then
    If ADOprimaryrs Is Nothing Then
    Else
        ADOprimaryrs.Close
        Set ADOprimaryrs = Nothing
    End If
    Set ADOprimaryrs = New ADODB.Recordset
    ADOprimaryrs.Open "SELECT [ID],[11] FROM [Letter] WHERE [ID]=" & tvwDB.SelectedItem.Tag, db, adOpenKeyset, adLockOptimistic, adCmdText
        With ADOprimaryrs
            If .RecordCount > 0 Then
                If IsNull(![11]) Then
                    ![11] = " "
                    .Update
                End If
                RTF.Text = ![11]
                'Set RTF.DataSource = ADOprimaryRS
                cmdUpdate.Caption = "Update - " & tvwDB.SelectedItem.Text
            End If
        End With
End If
ShowStatus False

End Sub
