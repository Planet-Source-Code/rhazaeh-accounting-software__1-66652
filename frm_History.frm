VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frm_History 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Loading History"
   ClientHeight    =   6330
   ClientLeft      =   4350
   ClientTop       =   3540
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   6495
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Ok"
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   5880
      Width           =   1575
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   5880
      Width           =   1575
   End
   Begin ComctlLib.ListView lstView 
      Height          =   5295
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   9340
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   5880
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_History.frx":0000
            Key             =   "tbs"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_History.frx":031A
            Key             =   "No"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Loading History"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   6495
   End
End
Attribute VB_Name = "frm_History"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TemptConnection As String
Dim TempgblADOProvider As String
Dim dbTemp As ADODB.Connection
Dim ADOprimaryrs As ADODB.Recordset
Dim WhichType As Integer

Public Sub WhichToExpose(UserInput As Integer)
      If UserInput > 1 Then
        UserInput = 1
      ElseIf UserInput < 0 Then
        UserInput = 0
      End If
      
      WhichType = UserInput
      LoadForm
      frm_History.Show vbModal
End Sub

Private Sub cmdLoad_Click()
    Select Case cmdLoad.Caption
    Case "Load/Quit"
        If lstView.ListItems(lstView.SelectedItem.Index).Text = "May Be in Mars" Then
            Beep
            MsgBox "Houston we got a PROBLEM", vbCritical, ""
        Else
            LoadDataBase lstView.ListItems(lstView.SelectedItem.Index).SubItems(1)
        End If
    Case "Clear/Quit"
        dbTemp.Execute "Delete * From [Error Log]"
        Unload Me
    End Select
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub LoadForm()
On Error GoTo FormErr
      
    If WhichType = 1 Then
        'TemptConnection = App.Path & "\properties.nos"
        TempgblADOProvider = gblBasicADOProvider
        Label1.Caption = "Loading History"
        Me.Caption = "Loading History"
        cmdLoad.Caption = "Load/Quit"
    ElseIf WhichType = 0 Then
        'TemptConnection = App.Path & "\properties.nos"
        TempgblADOProvider = gblADOProvider
        Label1.Caption = "Error Log"
        Me.Caption = "Error Log"
        cmdLoad.Caption = "Clear/Quit"
    End If
    
      Set dbTemp = New ADODB.Connection
      dbTemp.CursorLocation = adUseClient
      dbTemp.Open TempgblADOProvider
      
      LoadListView

Exit Sub
FormErr:
    MsgBox "Missing Database. Please contact the supplier", vbCritical, "Error"
    Unload Me
End Sub

Private Sub LoadListView()

Dim SQLstatement As String
Dim clmH As ColumnHeader
Dim SetItems As ListItem

      lstView.View = lvwReport
      
            
      Dim i As Integer
      Dim FileExist As String
      
      '
      If WhichType = 1 Then
        SQLstatement = "SELECT * FROM [Last Company]"
        Set ADOprimaryrs = New ADODB.Recordset
        ADOprimaryrs.Open SQLstatement, dbTemp, adOpenKeyset, adLockOptimistic, adCmdText
        
        Set clmH = lstView.ColumnHeaders.Add(, , "File", 1500)
        Set clmH = lstView.ColumnHeaders.Add(, , "Folder", 2000)
        Set clmH = lstView.ColumnHeaders.Add(, , "Date", 1000)
        Set clmH = lstView.ColumnHeaders.Add(, , "User", 1000)
        
        With ADOprimaryrs
        
        If .RecordCount = 0 Then
            Exit Sub
        End If
        
        .MoveFirst
        
            i = 0
            
            Do While ADOprimaryrs.EOF = False
                FileExist = Dir(![Last Company])
                If FileExist = "" Then
                    Set SetItems = lstView.ListItems.Add(, "Col" & i, "May Be in Mars", , "No")
                Else
                    Set SetItems = lstView.ListItems.Add(, "Col" & i, FileExist, , "tbs")
                End If
                SetItems.SubItems(1) = ![Last Company]
                SetItems.SubItems(2) = ![Date]
                SetItems.SubItems(3) = ![Last User]
                ADOprimaryrs.MoveNext
                i = i + 1
            Loop
        End With
        
      ElseIf WhichType = 0 Then
      
        SQLstatement = "SELECT * FROM [Error Log]"
        Set ADOprimaryrs = New ADODB.Recordset
        ADOprimaryrs.Open SQLstatement, dbTemp, adOpenKeyset, adLockOptimistic, adCmdText

        Set clmH = lstView.ColumnHeaders.Add(, , "Form Name", 1500)
        Set clmH = lstView.ColumnHeaders.Add(, , "Module Name", 2000)
        Set clmH = lstView.ColumnHeaders.Add(, , "Date", 1000) 'When
        Set clmH = lstView.ColumnHeaders.Add(, , "Error Code", 1000)
        Set clmH = lstView.ColumnHeaders.Add(, , "Error String", 1000)
        
        With ADOprimaryrs
        
        If .RecordCount = 0 Then
            Exit Sub
        End If
        
        .MoveFirst
        
        i = 0
        
            Do While ADOprimaryrs.EOF = False
                Set SetItems = lstView.ListItems.Add(, "Col" & i, ![Form Name], , "tbs")
                SetItems.SubItems(1) = ![Module Name]
                SetItems.SubItems(2) = ![When]
                SetItems.SubItems(3) = ![Error Code]
                SetItems.SubItems(3) = ![Error String]
                ADOprimaryrs.MoveNext
                i = i + 1
            Loop
         End With
      End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    ADOprimaryrs.CancelUpdate
    ADOprimaryrs.Close
    Set ADOprimaryrs = Nothing
    dbTemp.Close
    Set dbTemp = Nothing
    Set frm_History = Nothing
End Sub

Private Sub lstView_DblClick()

    If lstView.ListItems(lstView.SelectedItem.Index).Text = "May Be in Mars" Then
        MsgBox "Houston we got a PROBLEM", vbCritical, ""
    Else
        LoadDataBase lstView.ListItems(lstView.SelectedItem.Index).SubItems(1)
    End If
End Sub

Private Sub LoadDataBase(FileName As String)
ShowStatus True
    'MsgBox FileName
   
   fMainForm.CloseAllMDIChild
   DbConnectionString FileName
   SaveCompany
    If gblApplicationConnectString = App.Path & "\properties.nos" Then
        fMainForm.txtConnection = "There is no Working database"
    Else
        fMainForm.txtConnection = gblApplicationConnectString
    End If
'MenuStatus True, True
ShowStatus False
Unload Me
End Sub
