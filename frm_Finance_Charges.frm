VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_Finance_Charges 
   Caption         =   "Asses Finance Charges"
   ClientHeight    =   5040
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12030
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5040
   ScaleWidth      =   12030
   Begin VB.Frame frPrimary 
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   12015
      Begin VB.Frame Frame1 
         Height          =   4095
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   3135
         Begin VB.TextBox txtfields 
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
            Index           =   1
            Left            =   600
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   1200
            Width           =   1335
         End
         Begin VB.CommandButton cmdDate 
            Height          =   285
            Index           =   0
            Left            =   1560
            Picture         =   "frm_Finance_Charges.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox txtfields 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "MM/dd/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   600
            Width           =   1335
         End
         Begin VB.CommandButton cmdClose 
            Caption         =   "&Close"
            Height          =   315
            Left            =   2040
            Picture         =   "frm_Finance_Charges.frx":030A
            TabIndex        =   2
            Top             =   3600
            Width           =   975
         End
         Begin VB.CommandButton cmdPrint 
            Caption         =   "Prin&t"
            Height          =   795
            Left            =   120
            Picture         =   "frm_Finance_Charges.frx":0EE6
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   2640
            Width           =   975
         End
         Begin VB.CommandButton cmdALL 
            Caption         =   "&Show All"
            Height          =   315
            Left            =   1080
            TabIndex        =   4
            Top             =   3600
            Width           =   975
         End
         Begin VB.CommandButton cmdPost 
            Caption         =   "&Post"
            Height          =   795
            Left            =   2040
            Picture         =   "frm_Finance_Charges.frx":11F0
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   2640
            Width           =   975
         End
         Begin VB.CommandButton cmdRefresh 
            Caption         =   "&Refresh"
            Height          =   315
            Left            =   120
            TabIndex        =   6
            Top             =   3600
            Width           =   975
         End
         Begin VB.Label lbllabels 
            Caption         =   "Total:"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   11
            Top             =   1200
            Width           =   495
         End
         Begin VB.Label lbllabels 
            Caption         =   "Finance Charge Date:"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   8
            Top             =   360
            Width           =   1695
         End
      End
      Begin MSDataGridLib.DataGrid grdDataGrid 
         Height          =   3975
         Left            =   3360
         TabIndex        =   7
         Top             =   240
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   7011
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
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
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "Customer"
            Caption         =   "Customer"
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
            DataField       =   "Finance Charge"
            Caption         =   "Finance Charge"
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
            DataField       =   "Period 1"
            Caption         =   "Period 1"
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
            DataField       =   "Period 2"
            Caption         =   "Period 2"
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
         BeginProperty Column04 
            DataField       =   "Period 3"
            Caption         =   "Period 3"
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
         BeginProperty Column05 
            DataField       =   "Period 4"
            Caption         =   "Period 4"
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
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1244.976
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1244.976
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1244.976
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1244.976
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1244.976
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frm_Finance_Charges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim db As ADODB.Connection
Dim ADOprimaryrs As ADODB.Recordset
Dim CustP1Balance  As Currency
Dim CustP2Balance  As Currency
Dim CustP3Balance As Currency
Dim CustP4Balance As Currency
Dim CustTotalBalance As Currency

Private Sub OpenDB()

BuildTable

Set ADOprimaryrs = New ADODB.Recordset
ADOprimaryrs.Open "Select * From [Finance Charge Work]", db, adOpenKeyset, adLockOptimistic, adCmdText

  Set grdDataGrid.DataSource = ADOprimaryrs

End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDate_Click(Index As Integer)
    Menu_Calendar.WhoCallMe True, 1302
End Sub

Private Sub cmdPost_Click()
  'On Error GoTo cmdPost_Click_Error

  'Post this invoice to the general ledger
   Dim Success%

  'Force record save
  'DoCmd.RunMacro "Save Record"

  'DoCmd.Hourglass True
  ShowStatus True

  db.BeginTrans
    Success% = PostFinanceCharges()
    If Success% = False Then
      db.RollbackTrans
      MsgBox "Transaction NOT Posted."
    Else
      db.CommitTrans
      MsgBox "Transaction Posted."
      Call BuildTable
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
  Call ErrorLog("Finance Charges", "cmdPost_Click", Now, Err.Number, Err.Description, True, db)
  Resume Next
End Sub

Private Sub BuildTable()
Dim SQLstatement As String

  Set grdDataGrid.DataSource = Nothing
  
  'On Error GoTo BuildTable_Error

  db.Execute "DELETE * FROM [Finance Charge Work]"

  Dim rsCustomer As ADODB.Recordset
  Set rsCustomer = New ADODB.Recordset '
  rsCustomer.Open "SELECT * FROM [AR Customer] where [AR CUST Finance Charge YN] = TRUE", db, adOpenKeyset, adLockReadOnly, adCmdText

  'On Error Resume Next
  'rsCustomer.MoveFirst
  If rsCustomer.RecordCount = 0 Then
    'On Error GoTo BuildTable_Error
    MsgBox "There are no customers set up with finance charges!", , "Error"
    'DoCmd.Close A_FORM, "Finance Charges"
    Exit Sub
  End If

  'On Error GoTo BuildTable_Error

  'Get the finance charge
  Dim FinanceChargePercent#
  Dim MinBalance#
  Dim MinFinanceCharge#
  Dim FinanceCharge#
  
  FinanceChargePercent# = LookRecord("[SYS COM Monthly Charge]", "[SYS Company]", db)
  MinBalance# = LookRecord("[SYS COM Minimum Balance]", "[SYS Company]", db)
  MinFinanceCharge# = LookRecord("[SYS COM Minimum Finance Charge]", "[SYS Company]", db)

  'Dim rsFinance As ADODB.Recordset
  'Set rsFinance = New ADODB.Recordset '
  'rsFinance.Open "SELECT * FROM [Finance Charge Work]", db, adOpenKeyset, adLockOptimistic, adCmdText
  
  Do While Not rsCustomer.EOF
    'Get customers financial information
    Call GetCustomerFinancialsLocal(CStr(rsCustomer("AR CUST Customer ID")))
    If CustTotalBalance > MinBalance# Then
      FinanceCharge# = CustTotalBalance * (FinanceChargePercent# / 100)
      If FinanceCharge# > MinFinanceCharge# Then
        
        SQLstatement = "INSERT INTO [Finance Charge Work]"
        SQLstatement = SQLstatement & " ([Customer],[Period 1],[Period 2],[Period 3],[Period 4],[Finance Charge])"
        SQLstatement = SQLstatement & " VALUES ('" & rsCustomer("AR CUST Customer ID") & "'," & CustP1Balance & "," & _
        CustP2Balance & "," & CustP3Balance & "," & CustP4Balance & "," & FinanceCharge# & ")"
        db.Execute SQLstatement
        
        'rsFinance.AddNew
        '  rsFinance("Customer") = rsCustomer("AR CUST Customer ID")
        '  rsFinance("Period 1") = CustP1Balance
        '  rsFinance("Period 2") = CustP2Balance
        '  rsFinance("Period 3") = CustP3Balance
        '  rsFinance("Period 4") = CustP4Balance
        '  rsFinance("Finance Charge") = FinanceCharge#
        'rsFinance.Update
      End If
    End If
    rsCustomer.MoveNext
  Loop
  rsCustomer.Close
  Set rsCustomer = Nothing
  
  
  Exit Sub
BuildTable_Error:
   ErrorLog "Finance Charges", "BuildTable", Now, Err.Number, Err.Description, True, db
End Sub

Private Sub GetCustomerFinancialsLocal(CustomerID$)

  'On Error GoTo CustomerFinancial_Error

  'Compute financial period balances for this customer

  'Get company's financial period information
  Dim Days&
  Dim AgeBy%
  Dim Period1%
  Dim Period2%
  Dim Period3%
  Dim Period4%
  Dim Balance#
  Dim TransDate As Variant  '
  Dim TransType$

  Dim ChargeInterest%
  ChargeInterest% = LookRecord("[SYS COM Interest YN]", "[SYS Company]", db)

  CustP1Balance = 0
  CustP2Balance = 0
  CustP3Balance = 0
  CustP4Balance = 0
  CustTotalBalance = 0

  Dim rsCompany As ADODB.Recordset
  Set rsCompany = New ADODB.Recordset
  rsCompany.Open "SELECT [SYS COM Sales Period 1],[SYS COM Sales Period 2],[SYS COM Sales Period 3] FROM [SYS Company]", db, adOpenKeyset, adLockOptimistic, adCmdText

  rsCompany.MoveFirst
  Period1% = rsCompany("SYS COM Sales Period 1")
  Period2% = rsCompany("SYS COM Sales Period 2")
  Period3% = rsCompany("SYS COM Sales Period 3")
  Period4% = 90
  AgeBy% = 2 'IIf(IsNull(rsCompany("SYS COM Sales Age Invoices By")), 1, rsCompany("SYS COM Sales Age Invoices By"))
  '1 - Invoice Date  2 - Due Date

  'Go through AR Sales and get transactions for this customer
  '   with balances > 0
  
  Dim rsARPay As ADODB.Recordset
  
  Dim ADOsales As ADODB.Recordset
  Set ADOsales = New ADODB.Recordset
  'Debug.Print "SELECT [AR SALE Customer ID],[AR SALE Ext Document #],[AR SALE Date],[AR SALE Due Date] FROM [AR Sales] where [AR SALE Customer ID] = '" & CustomerID$ & "' AND [AR SALE Balance Due] > 0 AND [AR SALE Posted YN] = TRUE"
  ADOsales.Open "SELECT [AR SALE Document Type],[AR SALE Customer ID],[AR SALE Balance Due],[AR SALE Ext Document #],[AR SALE Date],[AR SALE Due Date] FROM [AR Sales] where [AR SALE Customer ID] = '" & CustomerID$ & "' AND [AR SALE Balance Due] > 0 AND [AR SALE Posted YN] = TRUE", db, adOpenKeyset, adLockOptimistic, adCmdText
  'On Error Resume Next
  If ADOsales.RecordCount > 0 Then
    ADOsales.MoveFirst
    'If Err = 0 Then
    Do While Not ADOsales.EOF
      
      'Get the balance
      Balance# = IIf(IsNull(ADOsales("AR SALE Balance Due")), 0, ADOsales("AR SALE Balance Due"))

      'Get Transaction Type to see if we sould increase or decrease the balance
      TransType$ = ADOsales("AR SALE Document Type")
      Select Case TransType$
      Case "Invoice", "Sales Memo", "Beginning Balance", "Finance Charge"
        If ChargeInterest% = False And TransType$ = "Finance Charge" Then Balance# = 0
      Case "Credit Memo"
        'Find payment information and back out unapplied amount
        
        Set rsARPay = New ADODB.Recordset
        rsARPay.Open "SELECT [AR PAY Unapplied Amount] FROM [AR Payment Header] " & _
        "WHERE [AR PAY Customer No]='" & ADOsales("AR SALE Customer ID") & _
        "' AND [AR PAY Check No]='CM " & ADOsales("AR SALE Ext Document #") & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
            'rsARPay.Index = "PrimaryKey"
            'rsARPay.Seek "=", ADOsales("AR SALE Customer ID"), ADOsales("AR SALE Ext Document #")
            If rsARPay.RecordCount > 0 Then
              Balance# = rsARPay("AR PAY Unapplied Amount") * -1
            End If
        rsARPay.Close
        Set rsARPay = Nothing
        
      Case "Return"
        Set rsARPay = New ADODB.Recordset
        rsARPay.Open "SELECT [AR PAY Unapplied Amount] FROM [AR Payment Header] " & _
        "WHERE [AR PAY Customer No]='" & ADOsales("AR SALE Customer ID") & _
        "' AND [AR PAY Check No]='RET " & ADOsales("AR SALE Ext Document #") & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
            'rsARPay.Index = "PrimaryKey"
            'rsARPay.Seek "=", ADOsales("AR SALE Customer ID"), ADOsales("AR SALE Ext Document #")
            If rsARPay.RecordCount > 0 Then
              Balance# = rsARPay("AR PAY Unapplied Amount") * -1
            End If
        rsARPay.Close
        Set rsARPay = Nothing
      Case Else
        Balance# = 0
      End Select

      'Get a date to age by
      If (AgeBy% = 1) Then 'Use Invoice Date
        TransDate = IIf(IsNull(ADOsales("AR SALE Date")), Format("1/1/" & Format(Now, "yyyy"), "Short Date"), ADOsales("AR SALE Date"))
      Else                 'Use Due Date
        TransDate = IIf(IsNull(ADOsales("AR SALE Due Date")), Format("1/1/" & Format(Now, "yyyy"), "Short Date"), ADOsales("AR SALE Due Date"))
      End If

      Days& = DateDiff("d", TransDate, txtfields(0).Text)
      Select Case Days&
      Case Is < 0
        'Don't use it
      Case 0 To Period1%
        CustP1Balance = CustP1Balance + Balance#
      Case Period1% To Period2%
        CustP2Balance = CustP2Balance + Balance#
      Case Period2% To Period3%
        CustP3Balance = CustP3Balance + Balance#
      Case Else
        CustP4Balance = CustP4Balance + Balance#
      End Select
      ADOsales.MoveNext
    Loop
  End If

  'Now do payments
  Set rsARPay = New ADODB.Recordset
  'rsARPay.Open "SELECT [AR PAY Unapplied Amount] FROM [AR Payment Header] "
  rsARPay.Open "SELECT [AR PAY UnApplied Amount],[AR PAY Type],[AR PAY Transaction Date] " & _
  "FROM [AR PAYMENT Header] where [AR PAY Customer No] = '" & CustomerID$ & "' AND [AR PAY NSF] = False AND [AR PAY Posted YN] = TRUE", db, adOpenKeyset, adLockOptimistic, adCmdText
  
  'On Error Resume Next
  'rsARPay.MoveFirst
  If rsARPay.RecordCount > 0 Then
    rsARPay.MoveFirst
    Do While Not rsARPay.EOF
      'Get the balance
      Balance# = IIf(IsNull(rsARPay("AR PAY UnApplied Amount")), 0, rsARPay("AR PAY UnApplied Amount"))
      'Back out payment if type is return or NSF
      Select Case rsARPay("AR PAY Type")
      Case "NSF"
        Balance# = Balance# * -1
      Case "Credit Memo", "Return"
        Balance# = 0
      End Select

      TransDate = IIf(IsNull(rsARPay("AR PAY Transaction Date")), Format("1/1/" & Format(Now, "yyyy"), "Short Date"), rsARPay("AR PAY Transaction Date"))
      
      Days& = DateDiff("d", TransDate, txtfields(0).Text)

      Select Case Days&
      Case Is < 0
        'Don't use it
      Case 0 To Period1%
        CustP1Balance = CustP1Balance - Balance#
      Case Period1% To Period2%
        CustP2Balance = CustP2Balance - Balance#
      Case Period2% To Period3%
        CustP3Balance = CustP3Balance - Balance#
      Case Else
        CustP4Balance = CustP4Balance - Balance#
      End Select
      rsARPay.MoveNext
    Loop
  End If
  
  rsARPay.Close
  Set rsARPay = Nothing

  CustTotalBalance = CustP1Balance + CustP2Balance + CustP3Balance + CustP4Balance

CustomerFinancials_Exit:
  Exit Sub

CustomerFinancial_Error:
  Call ErrorLog("Finance Charges", "GetCustomerFinancialsLocal", Now, Err.Number, Err.Description, True, db)
  Resume Next
  
End Sub
Private Function PostFinanceCharges() As Integer
Dim SQLstatement As String

  'On Error GoTo PostFinanceCharges_error

  Dim msg$
  Dim title$

  Dim rsCompany As ADODB.Recordset
  
  Set rsCompany = New ADODB.Recordset
  rsCompany.Open "SELECT [SYS COM GL Post By Date],[SYS COM Sales AR Acct]," & _
  "[SYS COM Finance Charge Acct] FROM [SYS Company]", db, adOpenKeyset, adLockOptimistic, adCmdText
  
  rsCompany.MoveFirst
  
  'Dim rsGLWorkDetail As Recordset
  'Set rsGLWorkDetail = db2.OpenRecordset("GL Work Detail")

  'Post by 1 - system date or 2 - Transaction date?
  Dim PostDate%
  PostDate% = rsCompany("SYS COM GL Post By Date")


  'Set Post Date
  Dim TranDate As Variant
  If PostDate% = 1 Then
    TranDate = DateValue(Format(Now, "Short Date"))
  Else
    TranDate = DateValue(txtfields(0).Text)
  End If
  
  'Verify period can be posted to
  'Send TranDate
  'Return PeriodToPost and PeriodClosed
  Dim PeriodToPost%
  Dim PeriodClosed%
  Call VerifyPeriod(TranDate, PeriodToPost%, PeriodClosed%, db)
  
  'Is period open?
  If PeriodClosed% = True Then
    MsgBox "Unable to post transaction to a Closed Period.", , "Post Invoice Error"
    GoTo UnableToPostChargesHere
  End If

  On Error GoTo PostFinanceCharges_error
  
  ' clear any GL Work records
  db.Execute "qryDeleteGLWorkDetail"

  'db2.OpenRecordset ("SELECT [AR CUST Financial Period 1],[AR CUST Highest Balance] FROM [AR Customer]")
  'rsCustomer.Index = "PrimaryKey"

  '--------------------------------------------------

  'Dim rsGLTrans As Recordset
  'Set rsGLTrans = db2.OpenRecordset("GL Transaction")

  ' write GL Transaction Header
  Dim refr$
  Dim desc$
  Dim NewNumber&
      
      SQLstatement = "INSERT INTO [GL Transaction]"
      SQLstatement = SQLstatement & " ([GL TRANS Document #],[GL TRANS Type],[GL TRANS Date],"
      SQLstatement = SQLstatement & " [GL TRANS Reference],[GL TRANS Amount],[GL TRANS Posted YN],"
      SQLstatement = SQLstatement & " [GL TRANS Description],[GL TRANS Source],[GL TRANS System Generated])"

  'rsGLTrans.AddNew
    'NewNumber& = rsGLTrans("GL TRANS Number")
    
    'rsGLTrans("GL TRANS Document #") = "FC " & Trim(CStr(NewNumber&))
    'rsGLTrans("GL TRANS Type") = "Finance Charges"
    
    ' gl post date
    Dim TempStr As String
    If PostDate% = 1 Then
      TempStr = Format(Now, "Short Date")
    Else
      TempStr = txtfields(0).Text
    End If
    
    SQLstatement = SQLstatement & " VALUES ('FC TEMP' & AppLoginName,'Finance Charges',#" & TempStr & "#,"
    
    refr$ = "Monthly Finance Charges"
    desc$ = "Monthly Finance Charges"
    SQLstatement = SQLstatement & "'" & refr$ & "'," & CCur(txtfields(1).Text) & ",1,"
    SQLstatement = SQLstatement & "'" & desc$ & "','FC TEMP' & AppLoginName,True)"
      'Debug.Print SQLstatement
      
    db.Execute SQLstatement
    
  Dim rsGLTrans As ADODB.Recordset
  Set rsGLTrans = New ADODB.Recordset
  rsGLTrans.Open "SELECT [GL TRANS Number] FROM [GL Transaction] WHERE [GL TRANS Document #]='" & "FC TEMP" & AppLoginName, db, adOpenKeyset, adLockReadOnly, adCmdText
      NewNumber& = rsGLTrans("GL TRANS Number")
      rsGLTrans("GL TRANS Document #") = "FC " & Trim(CStr(NewNumber&))
      rsGLTrans("GL TRANS Source") = "FC " & Trim(CStr(NewNumber&))
      rsGLTrans.Update
  rsGLTrans.Close
  Set rsGLTrans = Nothing
  '  rsGLTrans("GL TRANS Reference") = refr$
  '  rsGLTrans("GL TRANS Amount") = txtfields(1).Text
  '  rsGLTrans("GL TRANS Posted YN") = 1
  '  rsGLTrans("GL TRANS Description") = desc$
  '  rsGLTrans("GL TRANS Source") = "FC " & Trim(CStr(NewNumber&))
  '  rsGLTrans("GL TRANS System Generated") = True
  'rsGLTrans.Update
  ' write GL Transaction Detail
  
  ' GL Debit
  ' AR
  SQLstatement = "INSERT INTO [GL Work Detail]"
  SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
  SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & rsCompany("SYS COM Sales AR Acct") & "'," & CCur(txtfields(1).Text) & ",0)"
  db.Execute SQLstatement
  'rsGLWorkDetail.AddNew
  '  rsGLWorkDetail("GW TRANSD Number") = NewNumber&
  '  rsGLWorkDetail("GW TRANSD Account") = rsCompany("SYS COM Sales AR Acct")
  '  rsGLWorkDetail("GW TRANSD Debit Amount") = txtfields(1).Text
  '  rsGLWorkDetail("GW TRANSD Credit Amount") = 0
  '  rsGLWorkDetail("GW TRANSD Project") = ""
  'rsGLWorkDetail.Update

  Dim rsDetail As ADODB.Recordset
  Set rsDetail = New ADODB.Recordset
  rsDetail.Open "SELECT * FROM [Finance Charge Work] WHERE [Finance Charge] > 0", db, adOpenKeyset, adLockOptimistic, adCmdText
  rsDetail.MoveFirst

  Dim rsSales As ADODB.Recordset
  'Set rsSales = db2.OpenRecordset("AR Sales")

  'Dim rsTerms As Recordset
  'Set rsTerms = db2.OpenRecordset("LIST Payment Terms")
  'rsTerms.Index = "PrimaryKey"

  Dim PaymentTermsSales$
  Dim DueDate As Variant
  Dim DiscountDate As Variant
  Dim DiscountDays%
  Dim CurrentBalance@

  Dim rsCustomer As ADODB.Recordset
  
  Do While Not rsDetail.EOF
    ' update customer stats
    'rsCustomer.Seek "=", rsDetail("Customer")
    
    Set rsCustomer = New ADODB.Recordset '
    rsCustomer.Open "SELECT [AR CUST Financial Period 1],[AR CUST Highest Balance] FROM [AR Customer] where [AR CUST Customer ID] ='" & rsDetail("Customer") & "'", db, adOpenKeyset, adLockOptimistic, adCmdText

      'rsCustomer.Edit
      CurrentBalance@ = IIf(IsNull(rsCustomer("AR CUST Financial Period 1")), 0, rsCustomer("AR CUST Financial Period 1"))
      CurrentBalance@ = CurrentBalance@ + rsDetail("Finance Charge")
      rsCustomer("AR CUST Financial Period 1") = CurrentBalance@
      If CurrentBalance@ > IIf(IsNull(rsCustomer("AR CUST Highest Balance")), 0, rsCustomer("AR CUST Highest Balance")) Then rsCustomer("AR CUST Highest Balance") = CurrentBalance@
    
    rsCustomer.Update
    rsCustomer.Close
    Set rsCustomer = Nothing
 
    ' write a sales Finance Charge Record
    'rsSales.AddNew
      'rsSales("AR SALE Ext Document #") = "FC " & Trim(CStr(NewNumber&))
      'xxx 7/24/97 v8.3 1c
    '  rsSales("AR SALE Ext Document #") = "FC " & Trim(CStr(rsSales("AR SALE Document #")))
    '  rsSales("AR SALE Document Type") = "Finance Charge"
    '  rsSales("AR SALE Customer ID") = rsDetail("Customer")
    '  rsSales("AR SALE Date") = txtFields(0).Text
    '  rsSales("AR SALE Due Date") = txtFields(0).Text
    '  rsSales("AR SALE Total") = rsDetail("Finance Charge")
    '  rsSales("AR SALE Amount Paid") = 0
    '  rsSales("AR SALE Posted YN") = True
    '  rsSales("AR SALE Balance Due") = rsDetail("Finance Charge")
    'rsSales.Update
      SQLstatement = "INSERT INTO [AR Sales]"
      SQLstatement = SQLstatement & " ([AR SALE Ext Document #],[AR SALE Document Type]," & _
      "[AR SALE Customer ID],[AR SALE Date],[AR SALE Due Date],[AR SALE Total]," & _
      "[AR SALE Amount Paid],[AR SALE Posted YN],[AR SALE Balance Due])"
      
      SQLstatement = SQLstatement & " VALUES ('FC TEMP" & AppLoginName & ",'Finance Charge','" & _
      rsDetail("Customer") & "',#" & txtfields(0).Text & "#,#" & txtfields(0).Text & "#," & _
      rsDetail("Finance Charge") & ",0,True," & rsDetail("Finance Charge") & ")"
      db.Execute SQLstatement
      
      Set rsSales = New ADODB.Recordset
      rsSales.Open "SELECT [AR SALE Document #],[AR SALE Ext Document #] FROM [AR Sales] WHERE [AR SALE Ext Document #]='FC TEMP" & AppLoginName & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
        rsSales![AR SALE Ext Document #] = "<" & AppLoginName & Format(Now, "MMdd") & Right(Format(rsSales![AR SALE Document #] + 6000, "0000"), 4) & ">"
        rsSales.Update
      rsSales.Close
      Set rsSales = Nothing
    ' update GL
    '  Transaction #1

    ' write GL Transaction Detail
  
    '-----------------------------------------------------------------------
    ' Finance Charge
    '
    '                  Debit   Credit   Source
    '                  -----   ------   ------
    ' AR                 X              Pref - Sales
    ' Finance Charge Inc         X      Pref - Finance Charge Income
    '-----------------------------------------------------------------------
  
    ' Credits
    ' Finance Charge Income
  SQLstatement = "INSERT INTO [GL Work Detail]"
  SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
  SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & rsCompany("SYS COM Finance Charge Acct") & "',0," & rsDetail("Finance Charge") & ")"
  db.Execute SQLstatement
    'rsGLWorkDetail.AddNew
    '  rsGLWorkDetail("GW TRANSD Number") = NewNumber&
    '  rsGLWorkDetail("GW TRANSD Account") = DLookup("[SYS COM Finance Charge Acct]", "SYS Company")
    '  rsGLWorkDetail("GW TRANSD Debit Amount") = 0
    '  rsGLWorkDetail("GW TRANSD Credit Amount") = rsDetail("Finance Charge")
    '  rsGLWorkDetail("GW TRANSD Project") = ""
    'rsGLWorkDetail.Update

Skip_FC:
    rsDetail.MoveNext
  Loop
  
  rsCompany.Close
  Set rsCompany = Nothing
  
  rsDetail.Close
  Set rsDetail = Nothing
  
  ' post GL for these Finance Charges
  Dim Success%
  Success% = PostGLWorkDetail(TranDate, NewNumber&, db)
  If Success% = False Then
    MsgBox "An error occurred writing GL Transaction!"
    PostFinanceCharges% = False
    Exit Function
  End If

  PostFinanceCharges% = True

  Exit Function


UnableToPostChargesHere:
  PostFinanceCharges% = False

PostFinanceCharges_error:

  Call ErrorLog("Finance Charges", "PostFinanceCharges", Now, Err.Number, Err.Description, True, db)
  PostFinanceCharges% = False
  Exit Function
  'Resume Next
  
End Function

Private Sub Form_Load()
ShowStatus True
  Set db = New ADODB.Connection
  db.CursorLocation = adUseClient
  db.Open gblADOProvider
  
  txtfields(0).Text = FormatDate(Now)
  
  grdDataGrid.Columns(2).Caption = "0 to " & LookRecord("[SYS COM Sales Period 1]", "[SYS Company]", db) & " Days"
  grdDataGrid.Columns(3).Caption = LookRecord("[SYS COM Sales Period 1]", "[SYS Company]", db) + 1 & " To " & LookRecord("[SYS COM Sales Period 2]", "[SYS Company]", db) & " Days"
  grdDataGrid.Columns(4).Caption = LookRecord("[SYS COM Sales Period 2]", "[SYS Company]", db) + 1 & " To " & LookRecord("[SYS COM Sales Period 3]", "[SYS Company]", db) & " Days"
  grdDataGrid.Columns(5).Caption = "Over " & LookRecord("[SYS COM Sales Period 3]", "[SYS Company]", db) & " Days"
  OpenDB
ShowStatus False
End Sub

Private Sub Form_Unload(Cancel As Integer)
ADOprimaryrs.Close
Set ADOprimaryrs = Nothing

db.Close
Set db = Nothing

Set frm_Finance_Charges = Nothing
End Sub
