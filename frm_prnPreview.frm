VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#2.1#0"; "CRViewer.dll"
Begin VB.Form frm_prnPreview 
   Caption         =   "Form1"
   ClientHeight    =   5025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10605
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5025
   ScaleWidth      =   10605
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   7335
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   10335
      Begin CRVIEWERLibCtl.CRViewer CRViewer1 
         Height          =   3855
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   6255
         DisplayGroupTree=   0   'False
         DisplayToolbar  =   -1  'True
         EnableGroupTree =   -1  'True
         EnableNavigationControl=   -1  'True
         EnableStopButton=   -1  'True
         EnablePrintButton=   -1  'True
         EnableZoomControl=   -1  'True
         EnableCloseButton=   -1  'True
         EnableProgressControl=   -1  'True
         EnableSearchControl=   -1  'True
         EnableRefreshButton=   -1  'True
         EnableDrillDown =   -1  'True
         EnableAnimationControl=   -1  'True
         EnableSelectExpertControl=   0   'False
         EnableToolbar   =   -1  'True
         DisplayBorder   =   -1  'True
         DisplayTabs     =   -1  'True
         DisplayBackgroundEdge=   -1  'True
         SelectionFormula=   ""
         EnablePopupMenu =   -1  'True
      End
   End
   Begin VB.Label lblTop 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Report"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "frm_prnPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rpt As Report
Dim Rst As ADODB.Recordset
Dim db As ADODB.Connection
Dim ADOprimaryrs As ADODB.Recordset
Dim mNode As Node

Private Sub Form_Load()
ShowStatus True
If db Is Nothing Then
  Set db = New ADODB.Connection
  db.CursorLocation = adUseClient
  db.Open gblADOProvider
End If

  CRReportSourcw
  'GetTextColor Me
  
ShowStatus False
End Sub

Private Sub CRReportSourcw()
If rpt Is Nothing Then
Else
    CRViewer1.Visible = True
    CRViewer1.ReportSource = rpt
    CRViewer1.ViewReport
End If
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
  
  If Me.Width <= 4830 Then
     Me.Width = 4830
  End If
  If Me.Height <= 5430 Then
     Me.Height = 5430
  End If
SkipResize:
    
  Frame1.Width = Me.Width - 200
  Frame1.Height = Me.Height - 900
  Frame1.Left = (Me.ScaleWidth - Frame1.Width) / 2
  lblTop.Left = Frame1.Left
  lblTop.Width = Frame1.Width
  'CRViewer1.Top = 0
  'CRViewer1.Left = 0
  CRViewer1.Height = Frame1.Height - 300
  CRViewer1.Width = Frame1.Width - CRViewer1.Left - 100

End Sub

Public Sub Record(DocNo As String, FormName As String, DocType As String)
 'Pre: Index = [AP Order].[AP ORDER Ext Document #]
    ' open the connection to the data provider
  
  Set db = New ADODB.Connection
  db.CursorLocation = adUseClient
  db.Open gblADOProvider
    
    Dim sql As String
    Dim sqlCompany As String
    Dim sqlAROrder As String
    Dim sqlAROrderDetail As String
    Dim sqlARSales As String
    Dim sqlARSalesDetail As String
    Select Case FormName
    Case "frm_AR_Quote_Entry", "frm_AR_Order_Entry"
        If FormName = "frm_AR_Order_Entry" Then
            Set rpt = New crpt_prnOrder
        Else
            Set rpt = New crpt_prnQuote
        End If
    sqlCompany = "[SYS Company].[SYS COM Company Name],[SYS Company].[SYS COM Address 1]," & _
        "[SYS Company].[SYS COM Address 2],[SYS Company].[SYS COM City]," & _
        "[SYS Company].[SYS COM State],[SYS Company].[SYS COM Postal]," & _
        "[SYS Company].[SYS COM Country],[SYS Company].[SYS COM Phone]," & _
        "[SYS Company].[SYS COM Fax],[SYS Company].[SYS COM Web]"
    sqlAROrder = "[AR Order].[AR ORDER Ext Document #],[AR Order].[AR ORDER PO ID],[AR ORDER].[AR ORDER Customer ID]," & _
        "[AR Order].[AR ORDER Billing Customer],[AR Order].[AR ORDER Billing Address 1]," & _
        "[AR Order].[AR ORDER Billing Address 2],[AR Order].[AR ORDER Billing City]," & _
        "[AR Order].[AR ORDER Billing State],[AR Order].[AR ORDER Billing Postal]," & _
        "[AR Order].[AR ORDER Billing Country],[AR Order].[AR ORDER Shipping Customer]," & _
        "[AR Order].[AR ORDER Shipping Address 1],[AR Order].[AR ORDER Shipping Address 2]," & _
        "[AR Order].[AR ORDER Shipping City],[AR Order].[AR ORDER Shipping State]," & _
        "[AR Order].[AR ORDER Shipping Postal],[AR Order].[AR ORDER Shipping Country]," & _
        "[AR Order].[AR ORDER Date],[AR Order].[AR ORDER Ship Date]," & _
        "[AR Order].[AR ORDER Shipping Method],[AR Order].[AR ORDER Salesperson]," & _
        "[AR Order].[AR ORDER Payment Terms],[AR Order].[AR ORDER Payment Method]," & _
        "[AR Order].[AR ORDER Subtotal],[AR Order].[AR ORDER Discount Percent]," & _
        "[AR Order].[AR ORDER Discount Amount],[AR Order].[AR ORDER Tax Percent]," & _
        "[AR Order].[AR ORDER Sales Tax],[AR Order].[AR ORDER Freight]," & _
        "[AR Order].[AR ORDER Total],[AR Order].[AR ORDER Amount Paid]," & _
        "[AR Order].[AR ORDER Balance Due]"
    sqlAROrderDetail = "[AR Order Detail].[AR ORDERD Item Id]," & _
        "[AR Order Detail].[AR ORDERD Description],[AR Order Detail].[AR ORDERD Qty]," & _
        "[AR Order Detail].[AR ORDERD Qty To Invoice],[AR Order Detail].[AR ORDERD Unit Price]," & _
        "[AR Order Detail].[AR ORDERD Discount %],[AR Order Detail].[AR ORDERD Tax Rate]," & _
        "[AR Order Detail].[AR ORDERD Item Total]"
    sql = "select " & sqlCompany & "," & sqlAROrder & "," & sqlAROrderDetail & _
        " from [SYS Company],[AR Order] inner join [AR Order Detail] " & _
        "on [AR Order].[AR ORDER Document #]=[AR Order Detail].[AR ORDERD Document #]" & _
        "where [AR Order].[AR ORDER Ext Document #]='" & Trim(DocNo) & "'"
        
    Case "frm_AR_Sales_Entry", "frm_AR_Return_Entry"
        If FormName = "frm_AR_Sales_Entry" Then
            Set rpt = New crpt_prnSales
        Else
            Set rpt = New crpt_prnReturn
        End If
        sqlCompany = "[SYS Company].[SYS COM Company Name],[SYS Company].[SYS COM Address 1]," & _
        "[SYS Company].[SYS COM Address 2],[SYS Company].[SYS COM City]," & _
        "[SYS Company].[SYS COM State],[SYS Company].[SYS COM Postal]," & _
        "[SYS Company].[SYS COM Country],[SYS Company].[SYS COM Phone]," & _
        "[SYS Company].[SYS COM Fax],[SYS Company].[SYS COM Web]"
    sqlARSales = "[AR Sales].[AR SALE PO ID]," & _
        "[AR Sales].[AR SALE Ext Document #],[AR Sales].[AR SALE Quote Document #]," & _
        "[AR Sales].[AR SALE Billing Customer],[AR Sales].[AR SALE Billing Address 1]," & _
        "[AR Sales].[AR SALE Billing Address 2],[AR Sales].[AR SALE Billing City]," & _
        "[AR Sales].[AR SALE Billing State],[AR Sales].[AR SALE Billing Postal]," & _
        "[AR Sales].[AR SALE Billing Country],[AR Sales].[AR SALE Shipping Customer]," & _
        "[AR Sales].[AR SALE Shipping Address 1],[AR Sales].[AR SALE Shipping Address 2]," & _
        "[AR Sales].[AR SALE Shipping City],[AR Sales].[AR SALE Shipping State]," & _
        "[AR Sales].[AR SALE Shipping Postal],[AR Sales].[AR SALE Shipping Country]," & _
        "[AR Sales].[AR SALE Date],[AR Sales].[AR SALE Ship Date]," & _
        "[AR Sales].[AR SALE Shipping Method],[AR Sales].[AR SALE Salesperson]," & _
        "[AR Sales].[AR SALE Payment Terms],[AR Sales].[AR SALE Payment Method]," & _
        "[AR Sales].[AR SALE Subtotal],[AR Sales].[AR SALE Discount Percent]," & _
        "[AR Sales].[AR SALE Discount Amount],[AR Sales].[AR SALE Tax Percent]," & _
        "[AR Sales].[AR SALE Sales Tax],[AR Sales].[AR SALE Freight]," & _
        "[AR Sales].[AR SALE Total]"
    sqlARSalesDetail = "[AR Sales Detail].[AR SALED Item Id]," & _
        "[AR Sales Detail].[AR SALED Description],[AR Sales Detail].[AR SALED Qty]," & _
        "[AR Sales Detail].[AR SALED Units],[AR Sales Detail].[AR SALED Unit Price]," & _
        "[AR Sales Detail].[AR SALED Discount %],[AR Sales Detail].[AR SALED Tax Rate]," & _
        "[AR Sales Detail].[AR SALED Item Total]"
    sql = "select " & sqlCompany & "," & sqlARSales & "," & sqlARSalesDetail & _
        " from [SYS Company],[AR Sales]" & _
        " inner join [AR Sales Detail]" & _
        " on [AR Sales].[AR SALE Document #]=[AR Sales Detail].[AR SALED Document #]" & _
        " where [AR Sales].[AR SALE Ext Document #]='" & DocNo & "'"
    
    'Case "frm_AR_Return_Entry"
    'Set rpt = New crpt_prnReturn
    'sqlCompany = "[SYS Company].[SYS COM Company Name],[SYS Company].[SYS COM Address 1]," & _
        "[SYS Company].[SYS COM Address 2],[SYS Company].[SYS COM City]," & _
        "[SYS Company].[SYS COM State],[SYS Company].[SYS COM Postal]," & _
        "[SYS Company].[SYS COM Country],[SYS Company].[SYS COM Phone]," & _
        "[SYS Company].[SYS COM Fax]"
    'sqlARSales = "[AR Sales].[AR SALE Document Type],[AR Sales].[AR SALE Billing Customer]," & _
        "[AR Sales].[AR SALE Billing Address 1]," & _
        "[AR Sales].[AR SALE Billing Address 2],[AR Sales].[AR SALE Billing City]," & _
        "[AR Sales].[AR SALE Billing State],[AR Sales].[AR SALE Billing Postal]," & _
        "[AR Sales].[AR SALE Billing Country]," & _
        "[AR Sales].[AR SALE Date],[AR Sales].[AR SALE Ship Date]," & _
        "[AR Sales].[AR SALE Shipping Method],[AR Sales].[AR SALE Salesperson]," & _
        "[AR Sales].[AR SALE Payment Terms],[AR Sales].[AR SALE Payment Method]," & _
        "[AR Sales].[AR SALE Subtotal],[AR Sales].[AR SALE Discount Percent]," & _
        "[AR Sales].[AR SALE Discount Amount],[AR Sales].[AR SALE Tax Percent]," & _
        "[AR Sales].[AR SALE Sales Tax],[AR Sales].[AR SALE Freight]," & _
        "[AR Sales].[AR SALE Total]"
    'sqlARSalesDetail = "[AR Sales Detail].[AR SALED Item Id]," & _
        "[AR Sales Detail].[AR SALED Description],[AR Sales Detail].[AR SALED Qty]," & _
        "[AR Sales Detail].[AR SALED Units],[AR Sales Detail].[AR SALED Unit Price]," & _
        "[AR Sales Detail].[AR SALED Discount %],[AR Sales Detail].[AR SALED Tax Rate]," & _
        "[AR Sales Detail].[AR SALED Item Total]"
    'sql = "select " & sqlCompany & "," & sqlARSales & "," & sqlARSalesDetail & _
        " from [SYS Company],[AR Sales]" & _
        " inner join [AR Sales Detail]" & _
        " on [AR Sales].[AR SALE Document #]=[AR Sales Detail].[AR SALED Document #]" & _
        " where [AR Sales].[AR SALE Ext Document #]='" & DocNo & "'"

    Case "frm_AR_Sales_Memo_Entry", "frm_AR_Credit_Entry"
        If FormName = "frm_AR_Sales_Memo_Entry" Then
            Set rpt = New crpt_prnSMemo
        Else
            Set rpt = New crpt_prnCMemo
        End If
    sqlCompany = "[SYS Company].[SYS COM Company Name],[SYS Company].[SYS COM Address 1]," & _
        "[SYS Company].[SYS COM Address 2],[SYS Company].[SYS COM City]," & _
        "[SYS Company].[SYS COM State],[SYS Company].[SYS COM Postal]," & _
        "[SYS Company].[SYS COM Country],[SYS Company].[SYS COM Phone]," & _
        "[SYS Company].[SYS COM Fax],[SYS Company].[SYS COM Web]"
    sqlARSales = "[AR Sales].[AR SALE PO ID],[AR Sales].[AR SALE Ext Document #]," & _
        "[AR Sales].[AR SALE Quote Document #],[AR Sales].[AR SALE Billing Customer]," & _
        "[AR Sales].[AR SALE Billing Address 1]," & _
        "[AR Sales].[AR SALE Billing Address 2],[AR Sales].[AR SALE Billing City]," & _
        "[AR Sales].[AR SALE Billing State],[AR Sales].[AR SALE Billing Postal]," & _
        "[AR Sales].[AR SALE Billing Country]," & _
        "[AR Sales].[AR SALE Date],[AR Sales].[AR SALE Ship Date]," & _
        "[AR Sales].[AR SALE Shipping Method],[AR Sales].[AR SALE Salesperson]," & _
        "[AR Sales].[AR SALE Payment Terms],[AR Sales].[AR SALE Payment Method]," & _
        "[AR Sales].[AR SALE Subtotal],[AR Sales].[AR SALE Discount Percent]," & _
        "[AR Sales].[AR SALE Discount Amount],[AR Sales].[AR SALE Tax Percent]," & _
        "[AR Sales].[AR SALE Sales Tax],[AR Sales].[AR SALE Freight]," & _
        "[AR Sales].[AR SALE Total]"
    sqlARSalesDetail = "[AR Sales Detail].[AR SALED Posting Account]," & _
        "[AR Sales Detail].[AR SALED Description]," & _
        "[AR Sales Detail].[AR SALED Item Total],[AR Sales Detail].[AR SALED Project]"
    sql = "select " & sqlCompany & "," & sqlARSales & "," & sqlARSalesDetail & _
        " from [SYS Company],[AR Sales]" & _
        " inner join [AR Sales Detail]" & _
        " on [AR Sales].[AR SALE Document #]=[AR Sales Detail].[AR SALED Document #]" & _
        " where [AR Sales].[AR SALE Ext Document #]='" & DocNo & "'"
        
    Case "frm_AP_Purchase_Entry", "frm_AP_Receiving_Entry", "frm_AP_RMA_Entry"
        If FormName = "frm_AP_Purchase_Entry" Then
            Set rpt = New crpt_prnPurchase
        ElseIf FormName = "frm_AP_Receiving_Entry" Then
            Set rpt = New crpt_prnReceiving
        Else
            Set rpt = New crpt_prnRMA
        End If
    sqlCompany = "[SYS Company].[SYS COM Company Name],[SYS Company].[SYS COM Address 1]," & _
        "[SYS Company].[SYS COM Address 2],[SYS Company].[SYS COM City]," & _
        "[SYS Company].[SYS COM State],[SYS Company].[SYS COM Postal]," & _
        "[SYS Company].[SYS COM Country],[SYS Company].[SYS COM Phone]," & _
        "[SYS Company].[SYS COM Fax],[SYS Company].[SYS COM Web]"
    sqlARSales = "[AP Purchase].[AP PO Ext Document No],[AP Purchase].[AP PO Vendor Invoice No]," & _
        "[AP Purchase].[AP PO Vendor Name],[AP Purchase].[AP PO Address 1]," & _
        "[AP Purchase].[AP PO Address 2],[AP Purchase].[AP PO City]," & _
        "[AP Purchase].[AP PO State],[AP Purchase].[AP PO Postal]," & _
        "[AP Purchase].[AP PO Country]," & _
        "[AP Purchase].[AP PO Remit Name],[AP Purchase].[AP PO Remit Address 1]," & _
        "[AP Purchase].[AP PO Remit Address 2],[AP Purchase].[AP PO Remit City]," & _
        "[AP Purchase].[AP PO Remit State],[AP Purchase].[AP PO Remit Postal]," & _
        "[AP Purchase].[AP PO Remit Country]," & _
        "[AP Purchase].[AP PO Date],[AP Purchase].[AP PO Ship Method]," & _
        "[AP Purchase].[AP PO Date Requested],[AP Purchase].[AP PO Ordered By]," & _
        "[AP Purchase].[AP PO Payment Terms],[AP Purchase].[AP PO Due Date]," & _
        "[AP Purchase].[AP PO Subtotal],[AP Purchase].[AP PO Discount Percent]," & _
        "[AP Purchase].[AP PO Discount Amt],[AP Purchase].[AP PO Misc Percent]," & _
        "[AP Purchase].[AP PO Misc Charges],[AP Purchase].[AP PO Shipping]," & _
        "[AP Purchase].[AP PO Total Amount],[AP Purchase].[AP PO Amount Paid]," & _
        "[AP Purchase].[AP PO Balance Due]"
    sqlARSalesDetail = "[AP Purchase Detail].[AP POD Item ID]," & _
        "[AP Purchase Detail].[AP POD Description]," & _
        "[AP Purchase Detail].[AP POD Qty]," & _
        "[AP Purchase Detail].[AP POD Units],[AP Purchase Detail].[AP POD Unit Cost]," & _
        "[AP Purchase Detail].[AP POD Item Total]"
    sql = "select " & sqlCompany & "," & sqlARSales & "," & sqlARSalesDetail & _
        " from [SYS Company],[AP Purchase]" & _
        " inner join [AP Purchase Detail]" & _
        " on [AP Purchase].[AP PO Document No]=[AP Purchase Detail].[AP POD Document No]" & _
        " where [AP Purchase].[AP PO Ext Document No]='" & DocNo & "'"
        
    Case "frm_AP_Voucher_Entry", "frm_AP_Credit_Entry"
        If FormName = "frm_AP_Voucher_Entry" Then
            Set rpt = New crpt_prnVoucher
        Else
            Set rpt = New crpt_prnPCredit
        End If
    sqlCompany = "[SYS Company].[SYS COM Company Name],[SYS Company].[SYS COM Address 1]," & _
        "[SYS Company].[SYS COM Address 2],[SYS Company].[SYS COM City]," & _
        "[SYS Company].[SYS COM State],[SYS Company].[SYS COM Postal]," & _
        "[SYS Company].[SYS COM Country],[SYS Company].[SYS COM Phone]," & _
        "[SYS Company].[SYS COM Fax],[SYS Company].[SYS COM Web]"
    sqlARSales = "[AP Purchase].[AP PO Document Type],[AP Purchase].[AP PO Ext Document No],[AP Purchase].[AP PO Vendor Invoice No]," & _
        "[AP Purchase].[AP PO Vendor Name],[AP Purchase].[AP PO Address 1]," & _
        "[AP Purchase].[AP PO Address 2],[AP Purchase].[AP PO City]," & _
        "[AP Purchase].[AP PO State],[AP Purchase].[AP PO Postal]," & _
        "[AP Purchase].[AP PO Country]," & _
        "[AP Purchase].[AP PO Remit Name],[AP Purchase].[AP PO Remit Address 1]," & _
        "[AP Purchase].[AP PO Remit Address 2],[AP Purchase].[AP PO Remit City]," & _
        "[AP Purchase].[AP PO Remit State],[AP Purchase].[AP PO Remit Postal]," & _
        "[AP Purchase].[AP PO Remit Country]," & _
        "[AP Purchase].[AP PO Date],[AP Purchase].[AP PO Ship Method]," & _
        "[AP Purchase].[AP PO Date Requested],[AP Purchase].[AP PO Ordered By]," & _
        "[AP Purchase].[AP PO Payment Terms],[AP Purchase].[AP PO Due Date],[AP Purchase].[AP PO Payment Method]," & _
        "[AP Purchase].[AP PO Subtotal],[AP Purchase].[AP PO Discount Percent]," & _
        "[AP Purchase].[AP PO Discount Amt],[AP Purchase].[AP PO Misc Percent]," & _
        "[AP Purchase].[AP PO Misc Charges],[AP Purchase].[AP PO Shipping]," & _
        "[AP Purchase].[AP PO Total Amount]"
    sqlARSalesDetail = "[AP Purchase Detail].[AP POD Posting Account]," & _
        "[AP Purchase Detail].[AP POD Description]," & _
        "[AP Purchase Detail].[AP POD Project ID]," & _
        "[AP Purchase Detail].[AP POD Item Total]"
    sql = "select " & sqlCompany & "," & sqlARSales & "," & sqlARSalesDetail & _
        " from [SYS Company],[AP Purchase]" & _
        " inner join [AP Purchase Detail]" & _
        " on [AP Purchase].[AP PO Document No]=[AP Purchase Detail].[AP POD Document No]" & _
        " where [AP Purchase].[AP PO Ext Document No]='" & DocNo & "'"
    
    End Select
    'Debug.Print sql
    Set Rst = New ADODB.Recordset
    Rst.Open sql, db, adOpenDynamic, adLockReadOnly, adCmdText
    'Debug.Print sql
    'MsgBox Rst.RecordCount
    rpt.Database.SetDataSource Rst
    
    'Name the preview form accordingly
    'Select Case FormName
    'Case "frm_AR_Quote_Entry", "frm_AR_Order_Entry"
    '    Me.Caption = "Print Preview: " & DocType
    'Case "frm_AR_Sales_Entry", "frm_AR_Return_Entry", "frm_AR_Sales_Memo_Entry", _
    '    "frm_AR_Credit_Entry"
        Me.Caption = "Print Preview: " & DocType
    'End Select

End Sub

Public Sub Report(ReportType As String, ReportName As String)
ShowStatus True
Dim SQLFound As Boolean

If db Is Nothing Then
    Set db = New ADODB.Connection
    db.Open gblADOProvider
    db.CursorLocation = adUseClient
End If
    SQLFound = False
    Select Case ReportType
    
    'Accounting
    Case "Accounting"
        Select Case ReportName
        Case "Chart Of Accounts"
            'Set rpt = New crpt_Acctg_ChartOfAccounts
            SQLFound = True
        End Select
        
    'Accounts Receivable
    Case "Accounts Receivable"
        Select Case ReportName
        Case "Daily Cash Forecast"
            'Set rpt = New crpt_Acctg_DailyCashForecast
            SQLFound = True
        End Select
        
    'Bank
    'Case "Bank"
    '    Select Case ReportName
    '    Case "Bank List"
    '        Set Rst = New ADODB.Recordset
    '        Rst.Open "[rpt - Acctg - Bank List]", db, adOpenKeyset, adLockReadOnly, _
    '            adCmdStoredProc
    '        Set rpt = New crpt_Acctg_BankList
    '        rpt.Database.SetDataSource Rst
    '        SQLFound = True
    '    Case "Bank Register"
    '        Set Rst = New ADODB.Recordset
    '        Rst.Open "[rpt - Acctg - Bank Register]", db, adOpenKeyset, adLockReadOnly, _
    '            adCmdStoredProc
    '        Set rpt = New crpt_Acctg_BankRegister
    '        rpt.Database.SetDataSource Rst
    '        SQLFound = True
    '    Case "Bank Transactions"
            'Set rpt = New crpt_Bank_BankTransactions
    '        SQLFound = True
    '    End Select
    End Select
    
    If SQLFound = True Then
        CRReportSourcw
    End If
ShowStatus False
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Error GoTo FormErr
  ShowStatus True
      If Rst Is Nothing Then
      Else
        Rst.Close
        Set Rst = Nothing
        rpt.DiscardSavedData
        Set rpt = Nothing
      End If
      db.Close
      Set db = Nothing
  ShowStatus False
Exit Sub
FormErr:
  MsgBox Err.Description
  ShowStatus False
End Sub

