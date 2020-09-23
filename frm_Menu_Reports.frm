VERSION 5.00
Begin VB.Form frm_Menu_Reports 
   Caption         =   "Report Menu"
   ClientHeight    =   7575
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   10470
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7575
   ScaleWidth      =   10470
   Begin VB.Frame frPrimary 
      Height          =   7095
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   10455
   End
End
Attribute VB_Name = "frm_Menu_Reports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ADOprimaryrs As ADODB.Recordset
Attribute ADOprimaryrs.VB_VarHelpID = -1
Dim db As ADODB.Connection

Dim mNode As Node

Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean

Private Sub Form_Load()
Screen.MousePointer = vbHourglass
  Set db = New ADODB.Connection
  db.CursorLocation = adUseClient
  db.Open gblADOProvider

  tvwDB.Sorted = True
  mbDataChanged = False
Screen.MousePointer = vbNormal
End Sub

Private Sub OpenDB()
Dim ADOSecondRS As ADODB.Recordset
Dim i As Integer
Dim Relatif As String

  Set ADOprimaryrs = New ADODB.Recordset
  
  ADOprimaryrs.Open "SELECT DISTINCT [SYS Report Category] FROM [SYS Reports]", db, adOpenKeyset, adLockOptimistic, adCmdText
  'Set grdDataGrid.DataSource = ADOprimaryrs
  ADOprimaryrs.MoveFirst
    
    Set mNode = tvwDB.Nodes.Add()
        Relatif = "NoSecret Accounting Tech."
        mNode.Text = Relatif
        mNode.Tag = Relatif
        mNode.Key = Relatif
    
    Set mNode = tvwDB.Nodes.Add(Relatif, tvwChild)
        Relatif = "Report"
        mNode.Text = Relatif
        mNode.Tag = Relatif
        mNode.Key = Relatif
  
  i = 0
  Do While Not ADOprimaryrs.EOF
    i = i + 1
    Set mNode = tvwDB.Nodes.Add("Report", tvwChild)
        Relatif = ADOprimaryrs![SYS Report Category]
        mNode.Text = Relatif
        mNode.Tag = Relatif
        mNode.Key = Relatif
    
    Set ADOSecondRS = New ADODB.Recordset
    ADOSecondRS.Open "SELECT DISTINCT [SYS Report Name] FROM [SYS Reports] WHERE [SYS Report Category]='" & ADOprimaryrs![SYS Report Category] & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
    Do While Not ADOSecondRS.EOF
        Set mNode = tvwDB.Nodes.Add(Relatif, tvwChild)
           mNode.Text = ADOSecondRS![SYS Report Name]
           mNode.Tag = ADOSecondRS![SYS Report Name]
           mNode.Key = ADOSecondRS![SYS Report Name]
        ADOSecondRS.MoveNext
    Loop
    ADOprimaryrs.MoveNext
    ADOSecondRS.Close
    Set ADOSecondRS = Nothing
  Loop
  tvwDB.Nodes("NoSecret Accounting Tech.").Expanded = True
  tvwDB.Nodes("Report").Expanded = True
  ADOprimaryrs.MoveFirst
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
  
  Me.Width = 10560
  Me.Height = 7950
  
SkipResize:
  frPrimary.Left = (Me.ScaleWidth - frPrimary.Width) / 2
  lblTop.Left = frPrimary.Left
  lblTop.Width = frPrimary.Width
  frPrimary.Top = (Me.ScaleHeight - frPrimary.Height) / 2 + 230
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Error GoTo FormErr
  Screen.MousePointer = vbHourglass
      If ADOprimaryrs.RecordCount > 0 Then
        If ADOprimaryrs.EditMode <> 0 Then
          ADOprimaryrs.CancelUpdate
        End If
      End If
      ADOprimaryrs.Close
      Set ADOprimaryrs = Nothing
      db.Close
      Set db = Nothing
  Screen.MousePointer = vbDefault
Exit Sub
FormErr:
  MsgBox Err.Description
  Screen.MousePointer = vbDefault
End Sub
Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub grdDataGrid_DblClick()
'Display the report selected from the Grid Control
'Clicking on the report name in the grid makes that record active, all we need to is use the
'show property to display that report
Dim RptName As String
'RptName = grdDataGrid.Columns(1).CellText(grdDataGrid.Bookmark)
Select Case RptName
   Case "Forms Invoice"
     Rpt___Forms_Invoice.Show
   Case "Forms Purchase Order"
     Rpt___Forms_Purchase_Order.Show
   Case "Forms Quote"
     Rpt___Forms_Quote.Show
   Case "Forms Receiving"
     Rpt___Forms_Receiving.Show
   Case "Forms Return"
     Rpt___Forms_Return.Show
   Case "Forms Statement"
     Rpt___Forms_Statement.Show
   Case "Plain Paper Credit Memo"
     Rpt___Plain_Paper_Credit_Memo.Show
   Case "Plain Paper Invoice"
     Rpt___Plain_Paper_Invoice.Show
   Case "Plain Paper InvoiceStmt"
     Rpt___Plain_Paper_InvoiceStmt.Show
   Case "Plain Paper Order"
     Rpt___Plain_Paper_Order.Show
   Case "Plain Paper PO Credit Memo"
     Rpt___Plain_Paper_PO_Credit_Memo.Show
   Case "Plain Paper Purchase Order"
     Rpt___Plain_Paper_Purchase_Order.Show
   Case "Plain Paper Quote"
     Rpt___Plain_Paper_Quote.Show
   Case "Plain Paper Receiving"
     Rpt___Plain_Paper_Receiving.Show
   Case "Plain Paper Return"
     Rpt___Plain_Paper_Return.Show
   Case "Plain Paper Sales Memo"
     Rpt___Plain_Paper_Sales_Memo.Show
   Case "Plain Paper Statement"
     Rpt___Plain_Paper_Statement.Show
   Case "Plain Paper Voucher"
     Rpt___Plain_Paper_Voucher.Show
   Case Else
       'Report is in the other executable, IABReports.exe, Pass the report name, Connection string  Connection Name to
       ' te executable and have it display the report
End Select


End Sub

