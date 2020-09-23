Attribute VB_Name = "Sales_Module"

Function AverageDaysToPay!(CustomerID$, db As ADODB.Connection)

  'On Error Resume Next

  'Calculate the average days to pay for this customer

  Dim sql$
  Dim Days&
  Dim ARID&
  Dim AvgDays!
  Dim NumPays%
  Dim LASTDATE&
  Dim TransDate&
  Dim InvoiceID$
  
  'Dim db As ADODB.Connection
  'Set db = New ADODB.Connection
  'db.CursorLocation = adUseClient
  'db.Open gblADOProvider
  
  Dim dn As ADODB.Recordset
  Dim dn2 As ADODB.Recordset

  NumPays% = 0
  Set dn = New ADODB.Recordset
  dn.Open "SELECT * FROM [AR Sales] WHERE [AR SALE Customer ID] = '" & CustomerID$ & "' AND [AR SALE Balance Due] = 0 AND [AR SALE Posted YN] = TRUE", db, adOpenStatic, adLockOptimistic, adCmdText
    
  'On Error GoTo OuttaHere  'If no paid invoices, exit
  'dn.MoveFirst
  Do While Not dn.EOF
    'Calculate days to pay for this invoice
    InvoiceID$ = dn("AR SALE Document #")
    ARID& = dn("AR SALE Document #")
    'Find last payment for this AR OPEN ID
    sql$ = "SELECT DISTINCTROW [AR Payment Header].*, [AR Payment Invoice Cross Reference].[AR CROSS Payed ID] "
    sql$ = sql$ & "FROM [AR Payment Header] INNER JOIN [AR Payment Invoice Cross Reference] ON [AR Payment Header].[AR PAY ID] = [AR Payment Invoice Cross Reference].[AR CROSS Payment ID] "
    sql$ = sql$ & "WHERE [AR Payment Header].[AR PAY Posted YN] = True AND [AR Payment Invoice Cross Reference].[AR CROSS Payed ID]=" & ARID& & " ORDER BY [AR Payment Header].[AR PAY Transaction Date]"
    Set dn2 = New ADODB.Recordset
    dn2.Open sql$, db, adOpenStatic, adLockOptimistic, adCmdText
        
    'On Error Resume Next
    'Err = 0
    If dn2.RecordCount = 0 Then
      'No payments for this invoice
    Else
      dn2.MoveLast
      LASTDATE& = dn2("AR PAY Transaction Date")
      TransDate& = dn("AR SALE Date")
      Days& = Days& + (LASTDATE& - TransDate&)
      NumPays% = NumPays% + 1
    End If
    dn.MoveNext
    dn2.Close
  Loop
  
  If (NumPays% = 0) Then
    AvgDays! = 0
  Else
    AvgDays! = Int(Days& / NumPays%)
  End If

OuttaHere:
  AverageDaysToPay! = AvgDays!
  dn.Close
  Set dn = Nothing
  'dn2.Close
  Set dn2 = Nothing
  'db.Close
  'Set db = Nothing
  
  Exit Function

  dn.Close
  Set dn = Nothing
  dn2.Close
  Set dn2 = Nothing
  'db.Close
  'Set db = Nothing

End Function

Function BuildAgedReceivablesOld()

  'On Error GoTo BuildAgedReceivables_Error

  Dim db As ADODB.Connection
  Set db = New ADODB.Connection
  db.CursorLocation = adUseClient
  db.Open gblADOProvider
    
  Dim cmdtemp As ADODB.Recordset
  cmdtemp.Open "DELETE * FROM [Print Aged Receivables Work]", db, , , adCmdText
  cmdtemp.Close
  Set cmdtemp = Nothing
    
  Dim rsWork As ADODB.Recordset
  rsWork.Open "Print Aged Receivables Work", db, adOpenStatic, adLockOptimistic, adCmdTable
                                
  Dim CustomerID$

  Dim rsCompany As ADODB.Recordset
  rsCompany.Open "SYS Company", db, adOpenStatic, adLockOptimistic, adCmdTable

  Dim Period1%
  Dim Period2%
  Dim Period3%
  Dim AgeBy%

  rsCompany.MoveFirst
  Period1% = rsCompany("SYS COM Sales Period 1")
  Period2% = rsCompany("SYS COM Sales Period 2")
  Period3% = rsCompany("SYS COM Sales Period 3")
  AgeBy% = IIf(IsNull(rsCompany("SYS COM Sales Age Invoices By")), 1, rsCompany("SYS COM Sales Age Invoices By"))
  '1 - Invoice Date  2 - Due Date

  Dim CurrentAmt@
  Dim Period1Amt@
  Dim Period2Amt@
  Dim Period3Amt@
  Dim Period4Amt@
  Dim TotalAmount@
  Dim CurrentPeriod%
  Dim Balance@
  Dim Days&

  TotalAmount@ = 0
  Balance@ = 0

  Dim TransAmount@
  Dim TransDate
  Dim rsARSales As ADODB.Recordset
  Dim rsARCustomer As ADODB.Recordset
  rsARCustomer.Open "AR Customer", db, adOpenStatic, adLockOptimistic, adCmdTable
    
  Dim Order%
  Order% = 0
  
  gMessage$ = "Formatting Report..."
        
  rsARCustomer.MoveFirst
  Do While Not rsARCustomer.EOF
    CustomerID$ = rsARCustomer("AR CUST Customer ID")
    rsARSales.Open "SELECT * FROM [AR Sales] WHERE [AR SALE Customer ID] = '" & CustomerID$ & "' AND [AR SALE Document Type] in ('Invoice','Beginning Balance','Sales Memo','Finance Charge') AND [AR SALE Posted YN] = TRUE ORDER BY [AR SALE Date] Asc", db, adOpenStatic, adLockOptimistic, adCmdText
    
    'On Error Resume Next
    rsARSales.MoveFirst
    Do While Not rsARSales.EOF
      rsWork.AddNew
        rsWork("Customer ID") = CustomerID$
        rsWork("Order") = Order%
        rsWork("Transaction Type") = rsARSales("AR SALE Document Type")
        rsWork("Transaction ID") = rsARSales("AR SALE Ext Document #")
        rsWork("Transaction Description") = rsARSales("AR SALE Description")
        rsWork("Applied To") = ""
  
        'What bucket do we use?
        'Get a date to age by
        If (AgeBy% = 1) Then 'Use Invoice Date
          TransDate = IIf(IsNull(rsARSales("AR SALE Date")), Format("1/1/" & Format(Now, "yyyy"), "Short Date"), rsARSales("AR SALE Date"))
        Else                 'Use Due Date
          TransDate = IIf(IsNull(rsARSales("AR SALE Due Date")), Format("1/1/" & Format(Now, "yyyy"), "Short Date"), rsARSales("AR SALE Due Date"))
        End If
        
        rsWork("Transaction Date") = TransDate
  
        If rsARSales("AR SALE Document Type") = "Return" Then
          TransAmount@ = rsARSales("AR SALE Total") * -1
        Else
          TransAmount@ = rsARSales("AR SALE Total")
        End If
  
        Days& = DateDiff("d", TransDate, Now)
        Select Case Days&
        Case Is < 0
          CurrentPeriod% = 1
          rsWork("Period 1") = TransAmount@
        Case 0 To Period1%
          CurrentPeriod% = 1
          rsWork("Period 1") = TransAmount@
        Case Period1% To Period2%
          CurrentPeriod% = 2
          rsWork("Period 2") = TransAmount@
        Case Period2% To Period3%
          CurrentPeriod% = 3
          rsWork("Period 3") = TransAmount@
        Case Else
          CurrentPeriod% = 4
          rsWork("Period 4") = TransAmount@
        End Select
  
        rsWork("Balance") = TransAmount@
      rsWork.Update
  
      Order% = Order% + 1
  
      TotalAmount@ = TotalAmount@ + TransAmount@
  
      If rsARSales("AR SALE Document Type") = "Return" Then GoTo SkipAgedPayments
  
      'Now load all payments to this invoice
      Err = 0
      Dim rsPayments As ADODB.Recordset
      rsPayments.Open "SELECT * FROM [qryCustomerPayments] where [AR Payment Invoice Cross Reference].[AR CROSS Payed ID] = " & rsARSales("AR SALE Document #") & " AND [AR PAY Posted YN] = True", db, adOpenStatic, adLockOptimistic, adCmdText
  
      If (rs.BOF And rs.EOF) Then GoTo SkipAgedPayments
  
      rsPayments.MoveFirst
      If (rs.BOF And rs.EOF) Then GoTo SkipAgedPayments
      
      Do While Not rsPayments.EOF
        If rsPayments("AR CROSS Applied Amount") >= 0.01 Then
          rsWork.AddNew
            rsWork("Customer ID") = CustomerID$
            rsWork("Order") = Order%
            rsWork("Transaction Date") = rsPayments("AR PAY Transaction Date")
            rsWork("Transaction Type") = rsPayments("AR PAY Type")
            rsWork("Transaction ID") = rsPayments("AR PAY Check No")
            rsWork("Transaction Description") = "Applied to " & rsARSales("AR SALE Ext Document #")
            rsWork("Applied To") = rsARSales("AR SALE Ext Document #")

            Select Case CurrentPeriod%
            Case 1
              rsWork("Period 1") = rsPayments("AR CROSS Applied Amount") * -1
            Case 2
              rsWork("Period 2") = rsPayments("AR CROSS Applied Amount") * -1
            Case 3
              rsWork("Period 3") = rsPayments("AR CROSS Applied Amount") * -1
            Case 4
              rsWork("Period 4") = rsPayments("AR CROSS Applied Amount") * -1
            End Select
  
            rsWork("Balance") = rsPayments("AR CROSS Applied Amount") * -1
          rsWork.Update
          Order% = Order% + 1
    
      
          TotalAmount@ = TotalAmount@ - rsPayments("AR CROSS Applied Amount")
        End If
  
        If rsPayments("AR CROSS Discount Taken") >= 0.01 Then
          rsWork.AddNew
            rsWork("Customer ID") = CustomerID$
            rsWork("Order") = Order%
            rsWork("Transaction Date") = rsPayments("AR PAY Transaction Date")
            rsWork("Transaction Type") = "Discount"
            rsWork("Transaction ID") = rsPayments("AR PAY Check No")
            rsWork("Transaction Description") = "Applied to " & rsARSales("AR SALE Ext Document #")
            rsWork("Applied To") = rsARSales("AR SALE Ext Document #")
            
            Select Case CurrentPeriod%
            Case 1
              rsWork("Period 1") = rsPayments("AR CROSS Discount Taken") * -1
            Case 2
              rsWork("Period 2") = rsPayments("AR CROSS Discount Taken") * -1
            Case 3
              rsWork("Period 3") = rsPayments("AR CROSS Discount Taken") * -1
            Case 4
              rsWork("Period 4") = rsPayments("AR CROSS Discount Taken") * -1
            End Select
            
            rsWork("Balance") = rsPayments("AR CROSS Discount Taken") * -1
          rsWork.Update
          
          Order% = Order% + 1
  
          TotalAmount@ = TotalAmount@ - rsPayments("AR CROSS Discount Taken")
        End If
  
        If rsPayments("AR CROSS Write Off Amount") >= 0.01 Then
          rsWork.AddNew
            rsWork("Customer ID") = CustomerID$
            rsWork("Order") = Order%
            rsWork("Transaction Date") = rsPayments("AR PAY Transaction Date")
            rsWork("Transaction Type") = "Write Off"
            rsWork("Transaction ID") = rsPayments("AR PAY Check No")
            rsWork("Transaction Description") = "Applied to " & rsARSales("AR SALE Ext Document #")
            rsWork("Applied To") = rsARSales("AR SALE Ext Document #")
            
            Select Case CurrentPeriod%
            Case 1
              rsWork("Period 1") = rsPayments("AR CROSS Write Off Amount") * -1
            Case 2
              rsWork("Period 2") = rsPayments("AR CROSS Write Off Amount") * -1
            Case 3
              rsWork("Period 3") = rsPayments("AR CROSS Write Off Amount") * -1
            Case 4
              rsWork("Period 4") = rsPayments("AR CROSS Write Off Amount") * -1
            End Select
            
            rsWork("Balance") = rsPayments("AR CROSS Write Off Amount") * -1
          rsWork.Update
          
          Order% = Order% + 1
      
          TotalAmount@ = TotalAmount@ - rsPayments("AR CROSS Write Off Amount")
        End If
  
        rsPayments.MoveNext
      Loop
SkipAgedPayments:
      rsARSales.MoveNext
    Loop
  
    'Now add payment from this customer that are not applied
    Err = 0
    rsPayments.Close
    Dim rsPayment As ADODB.Recordset
    rsPayment.Open "SELECT * FROM [qryCustomerPayments2] where [AR Payment Header].[AR PAY Customer No] = '" & CustomerID$ & "' AND [AR Payment Header].[AR PAY Unapplied Amount] > 0 AND [AR PAY Posted YN] = TRUE", db, adOpenStatic, adLockOptimistic, adCmdText
  
    If (rs.BOF And rs.EOF) Then GoTo SkipAgedPayments2
    rsPayment.MoveFirst
    If (rs.BOF And rs.EOF) Then GoTo SkipAgedPayments2
    Do While Not rsPayment.EOF
      rsWork.AddNew
        rsWork("Customer ID") = CustomerID$
        rsWork("Order") = Order%
        rsWork("Transaction Date") = rsPayment("AR PAY Transaction Date")
        rsWork("Transaction Type") = rsPayment("AR PAY Type")
        rsWork("Transaction ID") = rsPayment("AR PAY Check No")
        rsWork("Transaction Description") = "Unapplied"
        rsWork("Applied To") = ""
        
        Days& = DateDiff("d", rsWork("Transaction Date"), Now)
        Select Case Days&
        Case Is < 0
          rsWork("Period 1") = rsPayment("AR PAY UnApplied Amount") * -1
        Case 0 To Period1%
          rsWork("Period 1") = rsPayment("AR PAY UnApplied Amount") * -1
        Case Period1% To Period2%
          rsWork("Period 2") = rsPayment("AR PAY UnApplied Amount") * -1
        Case Period2% To Period3%
          rsWork("Period 3") = rsPayment("AR PAY UnApplied Amount") * -1
        Case Else
          rsWork("Period 4") = rsPayment("AR PAY UnApplied Amount") * -1
        End Select
        
        Balance@ = Balance@ - rsPayment("AR PAY UnApplied Amount")
        rsWork("Balance") = rsPayment("AR PAY UnApplied Amount") * -1
        
      rsWork.Update
      Order% = Order% + 1
    
      rsPayment.MoveNext
    Loop
    
SkipAgedPayments2:
    rsARCustomer.MoveNext
  Loop

  rsWork.Close
  Set rsWork = Nothing
  rsPayments.Close
  Set rsPayments = Nothing
  rsPayment.Close
  Set rsPayment = Nothing
  rsARCustomer.Close
  Set rsARCustomer = Nothing
  rsARSales.Close
  Set rsARSales = Nothing
  rsCompany.Close
  Set rsCompany = Nothing
  db.Close
  Set db = Nothing
  
  Exit Function
BuildAgedReceivables_Error:
  Call ErrorLog("Sales Module", "BuildAgedReceivables", Now, Err.Number, Err.Description, True, db)
  Resume Next

  rsWork.Close
  Set rsWork = Nothing
  rsPayments.Close
  Set rsPayments = Nothing
  rsPayment.Close
  Set rsPayment = Nothing
  rsARCustomer.Close
  Set rsARCustomer = Nothing
  rsARSales.Close
  Set rsARSales = Nothing
  rsCompany.Close
  Set rsCompany = Nothing
  db.Close
  Set db = Nothing

End Function

Function BuildAgedReceivables()
  
  'On Error GoTo BuildAgedReceivables_Error

  Dim db As ADODB.Connection
  Set db = New ADODB.Connection
  db.CursorLocation = adUseClient
  db.Open gblADOProvider
  
  Dim rstInfo As ADODB.Recordset
  
  rstInfo.Open "exp_qryGetPeriod_n_AgingInfo", db, adOpenStatic, adLockOptimistic, adCmdTable
  intPeriod1 = rstInfo("SYS COM Sales Period 1")
  intPeriod2 = rstInfo("SYS COM Sales Period 2")
  intPeriod3 = rstInfo("SYS COM Sales Period 3")
  
  Dim cmdtemp As ADODB.Recordset
  cmdtemp.Open "DELETE * FROM [Print Aged Receivables Work]", db, , , adCmdText
  cmdtemp.Close
  Set cmdtemp = Nothing
  
  'Find Proper Period for Amounts
  'Pull Discounts Applied to Aged Sales
  Dim expqryProcessARRDiscounts As String
  expqryProcessARRDiscounts = "INSERT INTO [Print Aged Receivables Work] ( [Customer ID], [Transaction Type], [Transaction Date], Balance, [Transaction ID], [Transaction Description], [Applied To], Period, [Order] ) SELECT DISTINCTROW [qryCustomerPayments].[AR PAY Customer No] AS [Customer ID], 'Discount' AS [Transaction Type], [qryCustomerPayments].[AR PAY Transaction Date]"
  expqryProcessARRDiscounts = expqryProcessARRDiscounts & " AS [Transaction Date], [AR CROSS Discount Taken]*-1 AS Balance, [qryCustomerPayments].[AR PAY Check No] AS [Transaction ID], 'Applied to ' & [AR SALE Ext Document #] AS [Transaction Description], [exp - qryAge Sales].[AR SALE Ext Document #] AS [Applied To], [exp - qryAge Sales].Period AS Expr1, [exp - qryAge Sales].[AR SALE Document #] AS Expr2 "
  expqryProcessARRDiscounts = expqryProcessARRDiscounts & " FROM qryCustomerPayments INNER JOIN [exp - qryAge Sales] ON [qryCustomerPayments].[AR CROSS Payed ID] = [exp - qryAge Sales].[AR SALE Document #] WHERE ((([qryCustomerPayments].[AR CROSS Discount Taken])>=0.01))"
  Dim cmdtemp1 As ADODB.Recordset
  cmdtemp1.Open expqryProcessARRDiscounts, db, , , adCmdText
  cmdtemp1.Close
  Set cmdtemp1 = Nothing
  
  'Pull Payments on Aged Sales
  Dim expqryProcessARRPayments As String
  expqryProcessARRPayments = "INSERT INTO [Print Aged Receivables Work] ( [Customer ID], [Transaction Type], [Transaction Date], Balance, [Transaction ID], [Transaction Description], [Applied To], Period, [Order] ) SELECT DISTINCTROW [qryCustomerPayments].[AR PAY Customer No] AS [Customer ID], [qryCustomerPayments].[AR PAY Type]"
  expqryProcessARRPayments = expqryProcessARRPayments & " AS [Transaction Type], [qryCustomerPayments].[AR PAY Transaction Date] AS [Transaction Date], [AR CROSS Applied Amount]*-1 AS Balance, [qryCustomerPayments].[AR PAY Check No] AS [Transaction ID], 'Applied to ' & [AR SALE Ext Document #] AS [Transaction Description], [exp - qryAge Sales].[AR SALE Ext Document #]"
  expqryProcessARRPayments = expqryProcessARRPayments & " AS [Applied To], [exp - qryAge Sales].Period AS Expr1, [exp - qryAge Sales].[AR SALE Document #] AS Expr2 FROM [exp - qryAge Sales] INNER JOIN qryCustomerPayments ON [exp - qryAge Sales].[AR SALE Document #] = [qryCustomerPayments].[AR CROSS Payed ID] WHERE ((([qryCustomerPayments].[AR CROSS Applied Amount])>=0.01))"
  Dim cmdtemp2 As ADODB.Recordset
  cmdtemp2.Open expqryProcessARRPayments, db, , , adCmdText
  cmdtemp2.Close
  Set cmdtemp2 = Nothing
 
  'Pull Write Offs on Aged Sales
  Dim expqryProcessARRWriteOffs As String
  expqryProcessARRWriteOffs = "INSERT INTO [Print Aged Receivables Work] ( [Customer ID], [Transaction Type], [Transaction Date], Balance, [Transaction Description], [Applied To], Period, [Order], [Transaction ID] ) SELECT DISTINCTROW [qryCustomerPayments].[AR PAY Customer No] AS [Customer ID], 'Write Off' AS [Transaction Type], [qryCustomerPayments].[AR PAY Transaction Date]"
  expqryProcessARRWriteOffs = expqryProcessARRWriteOffs & " AS [Transaction Date], [AR CROSS Write Off Amount]*-1 AS Balance, 'Applied to ' & [AR SALE Ext Document #] AS [Transaction Description], [exp - qryAge Sales].[AR SALE Ext Document #] AS [Applied To], [exp - qryAge Sales].Period AS Expr1, [exp - qryAge Sales].[AR SALE Document #] AS Expr2, [qryCustomerPayments].[AR PAY Check No] AS [Transaction ID] FROM [exp - qryAge Sales]"
  expqryProcessARRWriteOffs = expqryProcessARRWriteOffs & " INNER JOIN qryCustomerPayments ON [exp - qryAge Sales].[AR SALE Document #] = [qryCustomerPayments].[AR CROSS Payed ID] WHERE ((([AR CROSS Write Off Amount]*-1)<>0))"
  Dim cmdtemp3 As ADODB.Recordset
  cmdtemp3.Open "expqryProcessARRWriteOffs", db, , , adCmdText
  cmdtemp3.Close
  Set cmdtemp3 = Nothing
  
  'Pull Unapplied Payments
  Dim expqryAppendUnappliedARtoWorkTable As String
  expqryAppendUnappliedARtoWorkTable = "INSERT INTO [Print Aged Receivables Work] ([Transaction Date], [Customer ID], Balance, Period, [Transaction Description], [Order], [Transaction Type], [Transaction ID] ) SELECT DISTINCTROW [exp - qry Unapplied Payments].[AR PAY Transaction Date], [exp - qry Unapplied Payments].[AR PAY Customer No], [AR PAY Unapplied Amount]*-1"
  expqryAppendUnappliedARtoWorkTable = expqryAppendUnappliedARtoWorkTable & " AS Amount, 1 AS Period, 'Unapplied' AS [Desc], -1 AS [Order], 'Unapplied' AS Desc2, [exp - qry Unapplied Payments].[AR PAY Check No] FROM [exp - qry Unapplied Payments]"
  Dim cmdtemp4 As ADODB.Recordset
  cmdtemp4.Open expqryAppendUnappliedARtoWorkTable, db, , , adCmdText
  cmdtemp4.Close
  Set cmdtemp4 = Nothing

  'Now add the actual sales
  Dim expqryAddAgedSalestoWorkTable As String
  expqryAddAgedSalestoWorkTable = "INSERT INTO [Print Aged Receivables Work] ( [Customer ID], Period, [Transaction ID], [Transaction Type], Balance, [Transaction Date], [Order] ) SELECT [exp - qryAge Sales].[AR SALE Customer ID] AS Expr1, [exp - qryAge Sales].Period AS Expr2, [exp - qryAge Sales].[AR SALE Ext Document #] AS Expr3, [exp - qryAge Sales].[AR SALE Document Type] AS Expr4, [exp - qryAge Sales].[Transaction Amount] AS Expr5, [exp - qryAge Sales].[AR SALE Date] AS Expr6, [exp - qryAge Sales].[AR SALE Document #] AS Expr7 FROM [exp - qryAge Sales]"
  Dim cmdtemp5 As ADODB.Recordset
  cmdtemp5.Open expqryAddAgedSalestoWorkTable, db, , , adCmdText
  cmdtemp5.Close
  Set cmdtemp5 = Nothing

  rstInfo.Close
  Set rstInfo = Nothing
  db.Close
  Set db = Nothing
  
  Exit Function
BuildAgedReceivables_Error:
  Call ErrorLog("Sales Module", "BuildAgedReceivables", Now, Err.Number, Err.Description, True, db)
  Resume Next

  rstInfo.Close
  Set rstInfo = Nothing
  db.Close
  Set db = Nothing

End Function

Function BuildStatement(CustomerID$, db As ADODB.Connection) As Integer

  'On Error GoTo BuildStatement_Error

  'Dim db As ADODB.Connection
  'Set db = New ADODB.Connection
  'db.CursorLocation = adUseClient
  'db.Open gblADOProvider
  
  Dim rsWork As ADODB.Recordset

  'Dim cmdtemp As ADODB.Recordset
  'Set cmdtemp = New ADODB.Recordset
  db.Execute "DELETE * FROM [Print Statement Work]", , adCmdText
  
  Set rsWork = New ADODB.Recordset
  rsWork.Open "[Print Statement Work]", db, adOpenStatic, adLockOptimistic, adCmdTable
                                
  Dim rsPayments As ADODB.Recordset

  Dim rsARSales As ADODB.Recordset
  Set rsARSales = New ADODB.Recordset
  rsARSales.Open "SELECT * FROM [AR Sales] WHERE [AR SALE Customer ID] = '" & CustomerID$ & "' AND [AR SALE Cleared] = False and [AR SALE Posted YN] = True AND [AR SALE Document Type] in ('Invoice','Beginning Balance','Sales Memo','Finance Charge') ORDER BY [AR SALE Date] Asc", db, adOpenStatic, adLockOptimistic, adCmdText
  
  Dim Order%
  Order% = 0

  Dim Current@
  Dim Period1@
  Dim Period2@
  Dim Period3@
  Dim Period4@
  Dim TotalAmount@
  Dim CurrentPeriod@
  Dim Balance@
  Dim HoldAmount@

  Current@ = 0
  Period1@ = 0
  Period2@ = 0
  Period3@ = 0
  Period4@ = 0
  TotalAmount@ = 0
  Balance@ = 0
  
  gMessage$ = "Formatting Statement"

  'On Error Resume Next
  rsARSales.MoveFirst
  Do While Not rsARSales.EOF
    rsWork.AddNew
      rsWork("Customer ID") = CustomerID$
      rsWork("Order") = Order%
      rsWork("Transaction Date") = rsARSales("AR SALE Date")
      rsWork("Transaction Type") = rsARSales("AR SALE Document Type") & ""
      rsWork("Transaction ID") = rsARSales("AR SALE Ext Document #") & ""
      If IsNull(rsWork("Transaction Description")) = False Then
        rsWork("Transaction Description") = rsARSales("AR SALE Description") & ""
      End If
      If rsARSales("AR SALE Document Type") = "Return" Then
        rsWork("Amount") = rsARSales("AR SALE Total") * -1
        Balance@ = Balance@ - rsARSales("AR SALE Total")
      Else
        rsWork("Amount") = rsARSales("AR SALE Total")
        Balance@ = Balance@ + rsARSales("AR SALE Total")
      End If
      HoldAmount@ = rsWork("Amount")
      rsWork("Balance") = Balance@
    rsWork.Update

    Order% = Order% + 1

    Select Case DateDiff("d", rsARSales("AR SALE Due Date"), Now)
    Case Is <= 0
      CurrentPeriod@ = 0
      Current@ = Current@ + rsARSales("AR SALE Total")
    Case 1 To 30
      CurrentPeriod@ = 1
      Period1@ = Period1@ + rsARSales("AR SALE Total")
    Case 31 To 60
      CurrentPeriod@ = 2
      Period2@ = Period2@ + rsARSales("AR SALE Total")
    Case 61 To 90
      CurrentPeriod@ = 3
      Period3@ = Period3@ + rsARSales("AR SALE Total")
    Case Is > 90
      CurrentPeriod@ = 4
      Period4@ = Period4@ + rsARSales("AR SALE Total")
    End Select

    TotalAmount@ = TotalAmount@ + HoldAmount@

    If rsARSales("AR SALE Document Type") = "Return" Then GoTo SkipPayments

    'Now load all payments to this invoice
    Err = 0
    Set rsPayments = New ADODB.Recordset
    rsPayments.Open "SELECT * FROM [qryCustomerPayments] where [AR Payment Invoice Cross Reference].[AR CROSS Payed ID] = " & rsARSales("AR SALE Document #") & "", db, adOpenStatic, adLockOptimistic, adCmdText
    
    If rsPayments.RecordCount Then GoTo SkipPayments

    'rsPayments.MoveFirst
    'If (rs.BOF And rs.EOF) Then GoTo SkipPayments
    
    Do While Not rsPayments.EOF
      If rsPayments("AR CROSS Applied Amount") >= 0.01 Then
        rsWork.AddNew
          rsWork("Customer ID") = CustomerID$
          rsWork("Order") = Order%
          rsWork("Transaction Date") = rsPayments("AR PAY Transaction Date")
          rsWork("Transaction Type") = rsPayments("AR PAY Type")
          rsWork("Transaction ID") = rsPayments("AR PAY Check No")
          rsWork("Transaction Description") = "Applied to " & rsARSales("AR SALE Ext Document #")
          rsWork("Amount") = rsPayments("AR CROSS Applied Amount") * -1
          Balance@ = Balance@ - rsPayments("AR CROSS Applied Amount")
          rsWork("Balance") = Balance@
        rsWork.Update
        Order% = Order% + 1
  
        Select Case CurrentPeriod@
        Case 0
          Current@ = Current@ - rsPayments("AR CROSS Applied Amount")
        Case 1
          Period1@ = Period1@ - rsPayments("AR CROSS Applied Amount")
        Case 2
          Period2@ = Period2@ - rsPayments("AR CROSS Applied Amount")
        Case 3
          Period3@ = Period3@ - rsPayments("AR CROSS Applied Amount")
        Case 4
          Period4@ = Period4@ - rsPayments("AR CROSS Applied Amount")
        End Select
    
        TotalAmount@ = TotalAmount@ - rsPayments("AR CROSS Applied Amount")
      End If

      If rsPayments("AR CROSS Discount Taken") >= 0.01 Then
        rsWork.AddNew
          rsWork("Customer ID") = CustomerID$
          rsWork("Order") = Order%
          rsWork("Transaction Date") = rsPayments("AR PAY Transaction Date")
          rsWork("Transaction Type") = "Discount"
          rsWork("Transaction ID") = rsPayments("AR PAY Check No")
          rsWork("Transaction Description") = "Applied to " & rsARSales("AR SALE Ext Document #")
          rsWork("Amount") = rsPayments("AR CROSS Discount Taken") * -1
          Balance@ = Balance@ - rsPayments("AR CROSS Discount Taken")
          rsWork("Balance") = Balance@
        rsWork.Update
        
        Order% = Order% + 1

        Select Case CurrentPeriod@
        Case 0
          Current@ = Current@ - rsPayments("AR CROSS Discount Taken")
        Case 1
          Period1@ = Period1@ - rsPayments("AR CROSS Discount Taken")
        Case 2
          Period2@ = Period2@ - rsPayments("AR CROSS Discount Taken")
        Case 3
          Period3@ = Period3@ - rsPayments("AR CROSS Discount Taken")
        Case 4
          Period4@ = Period4@ - rsPayments("AR CROSS Discount Taken")
        End Select
    
        TotalAmount@ = TotalAmount@ - rsPayments("AR CROSS Discount Taken")
      End If
      
      If rsPayments("AR CROSS Write Off Amount") >= 0.01 Then
        rsWork.AddNew
          rsWork("Customer ID") = CustomerID$
          rsWork("Order") = Order%
          rsWork("Transaction Date") = rsPayments("AR PAY Transaction Date")
          rsWork("Transaction Type") = "Write Off"
          rsWork("Transaction ID") = rsPayments("AR PAY Check No")
          rsWork("Transaction Description") = "Applied to " & rsARSales("AR SALE Ext Document #")
          rsWork("Amount") = rsPayments("AR CROSS Write Off Amount") * -1
          Balance@ = Balance@ - rsPayments("AR CROSS Write Off Amount")
          rsWork("Balance") = Balance@
        rsWork.Update
        
        Order% = Order% + 1

        Select Case CurrentPeriod@
        Case 0
          Current@ = Current@ - rsPayments("AR CROSS Write Off Amount")
        Case 1
          Period1@ = Period1@ - rsPayments("AR CROSS Write Off Amount")
        Case 2
          Period2@ = Period2@ - rsPayments("AR CROSS Write Off Amount")
        Case 3
          Period3@ = Period3@ - rsPayments("AR CROSS Write Off Amount")
        Case 4
          Period4@ = Period4@ - rsPayments("AR CROSS Write Off Amount")
        End Select
    
        TotalAmount@ = TotalAmount@ - rsPayments("AR CROSS Write Off Amount")
      End If

      rsPayments.MoveNext
    Loop

SkipPayments:
    rsARSales.MoveNext
  Loop

  'Now add payment from this customer that are not applied
  Err = 0
  Dim rsPayment As ADODB.Recordset
  Set rsPayment = New ADODB.Recordset
  rsPayment.Open "SELECT * FROM [qryCustomerPayments2] where [AR Payment Header].[AR PAY Customer No] = '" & CustomerID$ & "' AND [AR Payment Header].[AR PAY Unapplied Amount] > 0 ", db, adOpenStatic, adLockOptimistic, adCmdText

  If rsPayment.RecordCount = 0 Then GoTo SkipPayments2
  rsPayment.MoveFirst
  'If (rs.BOF And rs.EOF) Then GoTo SkipPayments2
  Do While Not rsPayment.EOF
    rsWork.AddNew
      rsWork("Customer ID") = CustomerID$
      rsWork("Order") = Order%
      rsWork("Transaction Date") = rsPayment("AR PAY Transaction Date")
      rsWork("Transaction Type") = rsPayment("AR PAY Type")
      rsWork("Transaction ID") = rsPayment("AR PAY Check No")
      rsWork("Transaction Description") = "Unapplied"
      rsWork("Amount") = rsPayment("AR PAY UnApplied Amount") * -1
      Balance@ = Balance@ - rsPayment("AR PAY UnApplied Amount")
      rsWork("Balance") = Balance@
    rsWork.Update
    Order% = Order% + 1
  
    'Back unapplied amounts out of current or should I age them?
    Current@ = Current@ - rsPayment("AR PAY Unapplied Amount")
    TotalAmount@ = TotalAmount@ - rsPayment("AR PAY Unapplied Amount")

    rsPayment.MoveNext
  Loop
  
SkipPayments2:

  'Write out balances
  rsWork.MoveFirst
  Do While Not rsWork.EOF
      rsWork("Current") = Current@
      rsWork("1-30 Days") = Period1@
      rsWork("31-60 Days") = Period2@
      rsWork("61-90 Days") = Period3@
      rsWork("Over 90 Days") = Period4@
      rsWork("Total") = TotalAmount@
    rsWork.Update
    rsWork.MoveNext
  Loop

  rsWork.Close
  Set rsWork = Nothing
  rsPayment.Close
  Set rsPayment = Nothing
  rsPayments.Close
  Set rsPayments = Nothing
  'rsCompany.Close
  Set rsCompany = Nothing
  rsARSales.Close
  Set rsARSales = Nothing
  db.Close
  Set db = Nothing
  
  Exit Function
BuildStatement_Error:
  Call ErrorLog("Sales Module", "BuildStatement", Now, Err.Number, Err.Description, True, db)
  Resume Next

  rsWork.Close
  Set rsWork = Nothing
  rsPayment.Close
  Set rsPayment = Nothing
  rsPayments.Close
  Set rsPayments = Nothing
  rsCompany.Close
  Set rsCompany = Nothing
  rsARSales.Close
  Set rsARSales = Nothing
  db.Close
  Set db = Nothing

End Function

Function CloneReceipt(DocumentKey&, AskForID%) As Integer

  'On Error GoTo CloneReceipt_Error

  Dim db As ADODB.Connection
  Set db = New ADODB.Connection
  db.CursorLocation = adUseClient
  db.Open gblADOProvider
  
  Dim rs As ADODB.Recordset
  Dim rs2 As ADODB.Recordset

  Dim rsRecur As ADODB.Recordset
  Set rsRecur = New ADODB.Recordset
  rsRecur.Open "[SYS Recurred]", db, adOpenStatic, adLockOptimistic, adCmdTable
  
  Set rs = New ADODB.Recordset
  rs.Open "[AR Payment Header]", db, adOpenStatic, adLockOptimistic, adCmdTable
  Set rs2 = New ADODB.Recordset
  rs2.Open "[AR Payment Header]", db, adOpenStatic, adLockOptimistic, adCmdTable

  'rs.Index = "AR PAY ID"
  'rs.Seek DocumentKey&
  rs.MoveFirst
  rs.Find "[AP PAY ID]=" & DocumentKey&   '<<<---use select statement

  Dim MyCounter&
  Dim MyCounter2&

  MyCounter& = DocumentKey&

  'On Error Resume Next
  'On Error GoTo CopyReceiptFailed
  db.BeginTrans

  Dim X%
  Dim count%
  count% = rs2.Fields.count
  rs2.AddNew
    'Add all current rs records
    'MyCounter2& = rs2("AR PAY ID")
    'For x% = 0 To count% - 1
    '  rs2(x%) = rs(x%)
    'Next x%
    For X% = 1 To count% - 1
    If IsNull(rs(X%)) = False Then
      If rs2(X%).Type = 202 Or rs2(X%).Type = 203 Then
        rs2(X%) = rs(X%) & ""
      Else
        rs2(X%) = rs(X%)
      End If
    End If
    Next X%

    'rs2("AR PAY ID") = MyCounter2&
    'Rename Check #
    If AskForID% = True Then
      gNewInvoice$ = InputBox("Enter new check #")
    Else
      'Create an invoice ID
      gNewInvoice$ = "xxx"
    End If
    If gNewInvoice$ = "" Then
      db.RollbackTrans
      CloneReceipt% = 1
      Exit Function
    End If
    rs2("AR PAY Check No") = gNewInvoice$
    rs2("AR PAY Posted YN") = False
    rs2("AR PAY Transaction Date") = Date
    rsRecur.AddNew
      rsRecur("Document Type") = "Cash Receipt"
      rsRecur("Document Number") = rs2("AR PAY Check No")
      rsRecur("Reference") = rs2("AR PAY Customer No")
      rsRecur("Amount") = rs2("AR PAY Amount")
    rsRecur.Update
  rs2.Update

  db.CommitTrans
  CloneReceipt% = True
  
  rs.Close
  Set rs = Nothing
  rs2.Close
  Set rs2 = Nothing
  rsRecur.Close
  Set rsRecur = Nothing
  db.Close
  Set db = Nothing
    
  Exit Function

CopyReceiptFailed:
  db.RollbackTrans
  CloneReceipt% = False
  rs.Close
  Set rs = Nothing
  rs2.Close
  Set rs2 = Nothing
  rsRecur.Close
  Set rsRecur = Nothing
  db.Close
  Set db = Nothing
  Exit Function

CloneReceipt_Error:
  Call ErrorLog("Sales Module", "CloneReceipt", Now, Err.Number, Err.Description, True, db)
  Resume Next
  
  rs.Close
  Set rs = Nothing
  rs2.Close
  Set rs2 = Nothing
  rsRecur.Close
  Set rsRecur = Nothing
  db.Close
  Set db = Nothing

End Function

Function CloneSales(DocumentKey&, db As ADODB.Connection) As Integer

  'On Error GoTo CloneSales_Error

  'Dim db As ADODB.Connection
  'Set db = New ADODB.Connection
  'db.CursorLocation = adUseServer
  'db.Open gblADOProvider
  
  Dim rs As ADODB.Recordset
  Dim rs2 As ADODB.Recordset

  Dim rsDetail As ADODB.Recordset
  Dim rsDetail2 As ADODB.Recordset
  
  Set rs = New ADODB.Recordset
  rs.Open "SELECT * FROM [AR Sales] WHERE [AR SALE Document #]=" & DocumentKey&, db, adOpenStatic, adLockOptimistic, adCmdText
  Set rs2 = New ADODB.Recordset
  'rs2.Open "[AR Sales]", db, adOpenStatic, adLockOptimistic, adCmdTable
  rs2.Open "SELECT * FROM [AR Sales] WHERE [AR SALE Document #]=" & DocumentKey&, db, adOpenStatic, adLockOptimistic, adCmdText

  'rs.Index = "PrimaryKey"
  'rs.Seek DocumentKey&
  'rs.MoveFirst
  'rs.Find "[AR SALE Document #]=" & DocumentKey&

  'Dim rsRecur As ADODB.Recordset
  'Set rsRecur = New ADODB.Recordset
  'rsRecur.Open "[SYS Recurred]", db, adOpenStatic, adLockOptimistic, adCmdTable
  
  Dim MyCounter&
  Dim MyCounter2&

  MyCounter& = DocumentKey&

  'On Error Resume Next
  'On Error GoTo 0
  db.BeginTrans

  Dim X%
  Dim count%
  count% = rs2.Fields.count
  rs2.AddNew
    'Add all current rs records
    'MyCounter2& = rs2("AR SALE Document #")
    'For x% = 0 To count% - 1
    '  rs2(x%) = rs(x%)
    'Next x%
    For X% = 1 To count% - 1
    '  rs2(X%) = rs(X%)
        If IsNull(rs(X%)) = False Then
          If rs2(X%).Type = 202 Or rs2(X%).Type = 203 Then
            rs2(X%) = rs(X%) & ""
          Else
            rs2(X%) = rs(X%)
          End If
        End If
    Next X%

    'rs2("AR SALE Document #") = MyCounter2&
    'Rename Ext Document #
    'If AskForInvoice% = True Then
    '  gNewInvoice$ = InputBox("Enter new invoice #")
    'Else
      'Create an invoice ID
    '  Dim rsSeek As ADODB.Recordset
    '  Set rsSeek = New ADODB.Recordset
    '  rsSeek.Open "[AR Sales]", db, adOpenStatic, adLockOptimistic, adCmdTable
      'rsSeek.Index = "Ext Document #"
      
    '  Dim Counter%
    '  Counter% = 1
    '  Dim Success%
    '  Success% = False
    '  Do While Not Success%
    '    gNewInvoice$ = rs2("AR SALE Ext Document #") & "-" & Trim(Str(Counter%))
        'Check if this newly created document exists
        'rsSeek.Seek gNewInvoice$
    '    rsSeek.MoveFirst
    '    rsSeek.Find "[AR SALE Ext Document #]='" & gNewInvoice$ & "'"
    '    If rsSeek.EOF Then
    '      Success% = True
    '    Else
    '      Success% = False
    '      Counter% = Counter% + 1
    '    End If
    '  Loop
    'End If
    'If gNewInvoice$ = "" Then
    '  db.Rollback
    '  CloneSales% = 1
    '  Exit Function
    'End If
    rs2("AR SALE Ext Document #") = "CloneInv" & AppLoginName
    If rs("AR SALE Document Type") = "Quote" Then
      rs2("AR SALE Document Type") = "Invoice"
    End If
    rs2("AR SALE Date") = Date
    rs2("AR SALE Recurring YN") = False
    rs2("AR SALE Recur Type") = "Never"
    rs2("AR SALE Posted YN") = False
    rs2("AR SALE Amount Paid") = 0
    rs2("AR SALE Check Number") = " "
    rs2("AR SALE Status") = "Open"
    'rsRecur.AddNew
    '  rsRecur("Document Type") = rs2("AR SALE Document Type")
    '  rsRecur("Document Number") = rs2("AR SALE Ext Document #")
    '  rsRecur("Reference") = rs2("AR SALE Customer ID")
    '  rsRecur("Amount") = rs2("AR SALE Total")
    'rsRecur.Update
  rs2.Update
  rs2.Close
  Set rs2 = Nothing
  
'  MyCounter2& = rs2("AR SALE Document #")
  
  Set rs2 = New ADODB.Recordset
  rs2.Open "SELECT [AR SALE Document #],[AR SALE Ext Document #] FROM [AR Sales] WHERE [AR SALE Ext Document #]='CloneInv" & AppLoginName & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
    MyCounter2& = rs2![AR SALE Document #]
    rs2![AR SALE Ext Document #] = "[" & AppLoginName & Format(Now, "MMdd") & Right(Format(MyCounter2&, "0000"), 4) & "]"
  rs2.Update
  rs2.Close
  Set rs2 = Nothing
  
  
  Dim DetailCounter&
  Set rsDetail = New ADODB.Recordset
  rsDetail.Open "SELECT * FROM [AR Sales Detail] where [AR SALED Document #] = " & MyCounter&, db, adOpenKeyset, adLockOptimistic, adCmdText
  'On Error Resume Next
  'Err = 0
  If rsDetail.RecordCount = 0 Then
    'No Detail
  Else
    'rsDetail.MoveLast
    rsDetail.MoveFirst
    'Create new detail record
    Set rsDetail2 = New ADODB.Recordset
    'rsDetail2.Open "[AR Sales Detail]", db, adOpenStatic, adLockOptimistic, adCmdTable
    rsDetail2.Open "SELECT * FROM [AR Sales Detail] where [AR SALED Document #] = " & MyCounter&, db, adOpenKeyset, adLockOptimistic, adCmdText
    Do While Not rsDetail.EOF
      count% = rsDetail.Fields.count
      rsDetail2.AddNew
        'DetailCounter& = rsDetail2("AR SALED ID")
        'For x% = 0 To count% - 1
        '  rsDetail2(x%) = rsDetail(x%)
        'Next x%
        For X% = 1 To count% - 1
        '  rs2(X%) = rs(X%)
            If IsNull(rsDetail2(X%)) = False Then
              If rsDetail2(X%).Type = 202 Or rsDetail2(X%).Type = 203 Then
                rsDetail2(X%) = rsDetail(X%) & ""
              Else
                rsDetail2(X%) = rsDetail(X%)
              End If
            End If
        Next X%
        'rsDetail2("AR SALED ID") = DetailCounter&
        rsDetail2("AR SALED Document #") = MyCounter2&
        'Add rest of detail records
      rsDetail2.Update
      rsDetail.MoveNext
    Loop
  End If

SkipDetail:
  
  db.CommitTrans
  CloneSales% = True
  rs.Close
  Set rs = Nothing
  rs2.Close
  Set rs2 = Nothing
  'rsRecur.Close
  'Set rsRecur = Nothing
  'rsSeek.Close
  'Set rsSeek = Nothing
  rsDetail.Close
  Set rsDetail = Nothing
  rsDetail2.Close
  Set rsDetail2 = Nothing
  'db.Close
  'Set db = Nothing

  Exit Function

CopyFailed:
  db.RollbackTrans
  CloneSales% = False
  rs.Close
  Set rs = Nothing
  rs2.Close
  Set rs2 = Nothing
  'rsRecur.Close
  'Set rsRecur = Nothing
  'rsSeek.Close
  'Set rsSeek = Nothing
  rsDetail.Close
  Set rsDetail = Nothing
  rsDetail2.Close
  Set rsDetail2 = Nothing
  'db.Close
  'Set db = Nothing

  Exit Function
  
CloneSales_Error:
  Call ErrorLog("Sales Module", "CloneSales", Now, Err.Number, Err.Description, True, db)
  Resume Next
  
  rs.Close
  Set rs = Nothing
  rs2.Close
  Set rs2 = Nothing
  'rsRecur.Close
  'Set rsRecur = Nothing
  'rsSeek.Close
  'Set rsSeek = Nothing
  rsDetail.Close
  Set rsDetail = Nothing
  rsDetail2.Close
  Set rsDetail2 = Nothing
  'db.Close
  'Set db = Nothing
  
End Function

Function CloneOrders(DocumentKey&, AskForInvoice%) As Integer

  'On Error GoTo CloneOrders_Error

  Dim db As ADODB.Connection
  Set db = New ADODB.Connection
  db.CursorLocation = adUseClient
  db.Open gblADOProvider

  Dim rs As ADODB.Recordset
  Dim rs2 As ADODB.Recordset

  Dim rsDetail As ADODB.Recordset
  Dim rsDetail2 As ADODB.Recordset
  Set rs = New ADODB.Recordset
  rs.Open "[AR Order]", db, adOpenStatic, adLockOptimistic, adCmdTable
  Set rs2 = New ADODB.Recordset
  rs2.Open "[AR Order]", db, adOpenStatic, adLockOptimistic, adCmdTable

'  rs.Index = "PrimaryKey"
  'rs.Seek DocumentKey&
  rs.MoveFirst
  rs.Find "[AR ORDER Document #]=" & DocumentKey&

  Dim rsRecur As ADODB.Recordset
  Set rsRecur = New ADODB.Recordset
  rsRecur.Open "[SYS Recurred]", db, adOpenStatic, adLockOptimistic, adCmdTable
  
  Dim MyCounter&
  Dim MyCounter2&

  MyCounter& = DocumentKey&

  'On Error Resume Next
  db.BeginTrans

  Dim X%
  Dim count%
  count% = rs2.Fields.count
  rs2.AddNew
    'Add all current rs records
    'MyCounter2& = rs2("AR ORDER Document #")
    'For x% = 0 To count% - 1
    '  rs2(x%) = rs(x%)
    'Next x%
    For X% = 1 To count% - 1
    If IsNull(rs(X%)) = False Then
      If rs2(X%).Type = 202 Or rs2(X%).Type = 203 Then
        rs2(X%) = rs(X%) & ""
      Else
        rs2(X%) = rs(X%)
      End If
    End If
    Next X%

    'rs2("AR ORDER Document #") = MyCounter2&
    'Rename Ext Document #
    If AskForInvoice% = True Then
      gNewInvoice$ = InputBox("Enter new order #")
    Else
      'Create an invoice ID
      Dim rsSeek As ADODB.Recordset
      Set rsSeek = New ADODB.Recordset
      rsSeek.Open "[AR Order]", db, adOpenStatic, adLockOptimistic, adCmdTable
      'rsSeek.Index = "Ext Document #"
      Dim Counter%
      Counter% = 1
      Dim Success%
      Success% = False
      Do While Not Success%
        gNewInvoice$ = rs2("AR ORDER Ext Document #") & "-" & Trim(Str(Counter%))
        'Check if this newly created document exists
        rsSeek.MoveFirst
        rsSeek.Find "[AR ORDER Ext Document #]='" & gNewInvoice$ & "'"
        If rsSeek.EOF Then
          Success% = True
        Else
          Success% = False
          Counter% = Counter% + 1
        End If
      Loop
    End If
    If gNewInvoice$ = "" Then
      db.RollbackTrans
      CloneOrders% = 1
      Exit Function
    End If
    rs2("AR ORDER Ext Document #") = gNewInvoice$
    If rs("AR ORDER Document Type") = "Quote" Then
      rs2("AR ORDER Document Type") = "Order"
    End If
    
    rs2("AR ORDER Date") = Date
    rs2("AR ORDER Recurring YN") = False
    rs2("AR ORDER Recur Type") = "Never"
    rs2("AR ORDER Posted YN") = False
    rs2("AR ORDER Amount Paid") = 0
    rs2("AR ORDER Check Number") = ""
    rs2("AR ORDER Invoiced") = False
    rsRecur.AddNew
      rsRecur("Document Type") = rs2("AR ORDER Document Type")
      rsRecur("Document Number") = rs2("AR ORDER Ext Document #")
      rsRecur("Reference") = rs2("AR ORDER Customer ID")
      rsRecur("Amount") = rs2("AR ORDER Total")
    rsRecur.Update
  rs2.Update

  Dim DetailCounter&
  Set rsDetail = New ADODB.Recordset
  rsDetail.Open "SELECT * FROM [AR Order Detail] where [AR ORDERD Document #] = " & MyCounter&, db, adOpenStatic, adLockOptimistic, adCmdText
  'On Error Resume Next
  Err = 0
  rsDetail.MoveLast
  rsDetail.MoveFirst
  If rsDetail.RecordCount = 0 Then
    'No Detail
  Else
    'Create new detail record
    Set rsDetail2 = New ADODB.Recordset
    rsDetail2.Open "[AR Order Detail]", db, adOpenStatic, adLockOptimistic, adCmdTable
    Do While Not rsDetail.EOF
      count% = rsDetail.Fields.count
      rsDetail2.AddNew
        'DetailCounter& = rsDetail2("AR ORDERD ID")
        'For x% = 0 To count% - 1
        '  rsDetail2(x%) = rsDetail(x%)
        'Next x%
        'rsDetail2("AR ORDERD ID") = DetailCounter&
        For X% = 1 To count% - 1
            If IsNull(rs(X%)) = False Then
              If rsDetail2(X%).Type = 202 Or rsDetail2(X%).Type = 203 Then
                rsDetail2(X%) = rsDetail(X%) & ""
              Else
                rsDetail2(X%) = rsDetail(X%)
              End If
            End If
        Next X%
        rsDetail2("AR ORDERD Document #") = MyCounter2&
        'Add rest of detail records
      rsDetail2.Update
      rsDetail.MoveNext
    Loop
  End If

SkipDetail:
  
  db.CommitTrans
  CloneOrders% = True
  rs.Close
  Set rs = Nothing
  rs2.Close
  Set rs2 = Nothing
  rsRecur.Close
  Set rsRecur = Nothing
  rsDetail.Close
  Set rsDetail = Nothing
  rsDetail2.Close
  Set rsDetail2 = Nothing
  rsSeek.Close
  Set rsSeek = Nothing
  db.Close
  Set db = Nothing

  Exit Function

CopyFailed:
  db.RollbackTrans
  CloneOrders% = False
    rs.Close
  Set rs = Nothing
  rs2.Close
  Set rs2 = Nothing
  rsRecur.Close
  Set rsRecur = Nothing
  rsDetail.Close
  Set rsDetail = Nothing
  rsDetail2.Close
  Set rsDetail2 = Nothing
  rsSeek.Close
  Set rsSeek = Nothing
  db.Close
  Set db = Nothing
  Exit Function
  
CloneOrders_Error:
  Call ErrorLog("Sales Module", "CloneOrders", Now, Err.Number, Err.Description, True, db)
  Resume Next
  
  rs.Close
  Set rs = Nothing
  rs2.Close
  Set rs2 = Nothing
  rsRecur.Close
  Set rsRecur = Nothing
  rsDetail.Close
  Set rsDetail = Nothing
  rsDetail2.Close
  Set rsDetail2 = Nothing
  rsSeek.Close
  Set rsSeek = Nothing
  db.Close
  Set db = Nothing
  
End Function

Function CreateBackOrder(DocumentKey&, AskForInvoice%) As String

  '''On Error GoTo CreateBackorder_Error

  Dim db As ADODB.Connection
  Set db = New ADODB.Connection
  db.CursorLocation = adUseServer
  db.Open gblADOProvider
  
  Dim rs As ADODB.Recordset
  Dim rs2 As ADODB.Recordset

  Dim rsDetail As ADODB.Recordset
  Dim rsDetail2 As ADODB.Recordset

  Set rs = New ADODB.Recordset
  Set rs2 = New ADODB.Recordset
  rs.Open "[AR Order]", db, adOpenStatic, adLockOptimistic, adCmdTable
  rs2.Open "[AR Order]", db, adOpenStatic, adLockOptimistic, adCmdTable
  
  'rs.Index = "PrimaryKey"
  rs.MoveFirst
  rs.Find "[AR ORDER Document #]=" & DocumentKey&
  If rs.EOF Then
    MsgBox "Data Not Found--call Razi", vbCritical, "Error"
  End If
  Dim rsRecur As ADODB.Recordset
  Set rsRecur = New ADODB.Recordset
  rsRecur.Open "[SYS Recurred]", db, adOpenStatic, adLockOptimistic, adCmdTable
  
  Dim MyCounter&
  Dim MyCounter2&

  MyCounter& = DocumentKey&

  ''On Error Resume Next
  db.BeginTrans

  Dim X%
  Dim count%
  count% = rs2.Fields.count
  rs2.AddNew
    'Add all current rs records
    'MyCounter2& = rs2("AR ORDER Document #")

    For X% = 1 To count% - 1
    If IsNull(rs(X%)) = False Then
      If rs2(X%).Type = 202 Or rs2(X%).Type = 203 Then
        rs2(X%) = rs(X%) & ""
      Else
        rs2(X%) = rs(X%)
      End If
    End If
    Next X%

    'rs2("AR ORDER Document #") = MyCounter2&   '<<<----------autonumber
    'Rename Ext Document #
    If AskForInvoice% = True Then
      gNewOrder$ = InputBox("Enter new order #")
    Else
      'Create an invoice ID
      Dim rsSeek As ADODB.Recordset
      Set rsSeek = New ADODB.Recordset
      rsSeek.Open "[AR Order]", db, adOpenStatic, adLockOptimistic, adCmdTable
      'rsSeek.Index = "Ext Document #"
      Dim Counter%
      Counter% = 1
      Dim Success%
      Success% = False
      Do While Not Success%
        gNewOrder$ = rs2("AR ORDER Ext Document #") & "-" & Trim(Str(Counter%))
        'Check if this newly created document exists
        rsSeek.MoveFirst
        rsSeek.Find "[AR ORDER Ext Document #]='" & gNewOrder$ & "'"
        If rsSeek.EOF Then
          Success% = True
        Else
          Success% = False
          Counter% = Counter% + 1
        End If
      Loop
    End If
    If gNewOrder$ = "" Then
      db.RollbackTrans
      CreateBackOrder = ""
      Exit Function
    End If
    rs2("AR ORDER Ext Document #") = gNewOrder$
    rs2("AR ORDER Document Type") = "Backorder"
    rs2("AR ORDER Invoiced") = False
    rs2("AR ORDER Date") = Date
    rs2("AR ORDER Recurring YN") = False
    rs2("AR ORDER Recur Type") = "Never"
    rs2("AR ORDER Posted YN") = False
    rs2("AR ORDER Amount Paid") = 0
    rs2("AR ORDER Check Number") = ""
    rs2![AR ORDER Subtotal] = 0
    rs2![AR ORDER Sales Tax] = 0
    rs2![AR ORDER Discount Amount] = 0
    rs2![AR ORDER Total] = 0
  rs2.Update
  'rs2.MoveLast
  'rs2.Requery
  'rs2.MoveLast
    
  MyCounter2& = rs2("AR ORDER Document #")

  Dim DetailCounter&

    Set rsDetail = New ADODB.Recordset
    rsDetail.Open "SELECT * FROM [AR Order Detail] where [AR ORDERD Document #] = " & MyCounter&, db, adOpenStatic, adLockOptimistic, adCmdText
  ''On Error Resume Next
  'Err = 0
  rsDetail.MoveLast
  ''debug.print rsDetail.RecordCount
  rsDetail.MoveFirst
  If (rs.BOF And rs.EOF) Then
    'No Detail
  Else
    Set rsDetail2 = New ADODB.Recordset
    rsDetail2.Open "[AR Order Detail]", db, adOpenStatic, adLockOptimistic, adCmdTable
    Do While Not rsDetail.EOF
      If rsDetail("AR ORDERD Qty") - rsDetail("AR ORDERD Qty To Invoice") > 0 Then
            count% = rsDetail.Fields.count
            rsDetail2.AddNew
              For X% = 1 To count% - 1
                  If IsNull(rsDetail(X%)) = False Then
                    If rsDetail2(X%).Type = 202 Or rsDetail2(X%).Type = 203 Then
                      rsDetail2(X%) = rsDetail(X%) & ""
                    Else
                      rsDetail2(X%) = rsDetail(X%)
                    End If
                  End If
              Next X%
              rsDetail2("AR ORDERD Qty") = rsDetail("AR ORDERD Qty") - rsDetail("AR ORDERD Qty To Invoice")
              rsDetail2("AR ORDERD Qty To Invoice") = 0
              'rsDetail2("AR ORDERD ID") = DetailCounter&
              rsDetail2("AR ORDERD Document #") = MyCounter2&
              rsDetail2![AR ORDERD Item Total] = 0
              'Add rest of detail records
            rsDetail2.Update
            'rsDetail2.Requery
            'rsDetail2.MoveLast
            'DetailCounter& = rsDetail2("AR ORDERD ID")
      End If
    rsDetail.MoveNext
    Loop
  End If

SkipDetail:
  
  db.CommitTrans
  CreateBackOrder = gNewOrder$
  rs.Close
  Set rs = Nothing
  rs2.Close
  Set rs2 = Nothing
  rsRecur.Close
  Set rsRecur = Nothing
  rsDetail.Close
  Set rsDetail = Nothing
  rsDetail2.Close
  Set rsDetail2 = Nothing
  rsSeek.Close
  Set rsSeek = Nothing
  db.Close
  Set db = Nothing

  Exit Function

CopyFailed:
  db.RollbackTrans
  CreateBackOrder = ""
  rs.Close
  Set rs = Nothing
  rs2.Close
  Set rs2 = Nothing
  rsRecur.Close
  Set rsRecur = Nothing
  rsDetail.Close
  Set rsDetail = Nothing
  rsDetail2.Close
  Set rsDetail2 = Nothing
  rsSeek.Close
  Set rsSeek = Nothing
  db.Close
  Set db = Nothing
  
  Exit Function
  
CreateBackorder_Error:
  Call ErrorLog("Sales Module", "CreateBackorder", Now, Err.Number, Err.Description, True, db)
  Resume Next
  
  rs.Close
  Set rs = Nothing
  rs2.Close
  Set rs2 = Nothing
  rsRecur.Close
  Set rsRecur = Nothing
  rsDetail.Close
  Set rsDetail = Nothing
  rsDetail2.Close
  Set rsDetail2 = Nothing
  rsSeek.Close
  Set rsSeek = Nothing
  db.Close
  Set db = Nothing


End Function


Function GetARInvoiceDiscount(InvoiceID&, vntInDate As Variant, FromInvoice%)

  ''On Error GoTo GetARInvoiceDiscount_Error

  Dim ReferNo&
  Dim vntInvoiceDate As Variant
  Dim PaymentTerms$
  Dim Discount#
  Dim DiscountDays#
  Dim Diff%
  Dim DiscountAmount@
  Dim vntDate As Variant

  vntDate = vntInDate
  
  Dim db As ADODB.Connection
  Set db = New ADODB.Connection
  db.CursorLocation = adUseClient
  db.Open gblADOProvider
  
  Dim rsSales As ADODB.Recordset
  Set rsSales = New ADODB.Recordset
  rsSales.Open "[AR Sales]", db, adOpenStatic, adLockOptimistic, adCmdTable

  Dim rsTerms As ADODB.Recordset
  Set rsTerms = New ADODB.Recordset
  rsTerms.Open "[List Payment Terms]", db, adOpenStatic, adLockOptimistic, adCmdTable

  'rsSales.Index = "PrimaryKey"
  rsSales.MoveFirst
  rsSales.Find "[AR SALE Document #]=" & InvoiceID&
  If rsSales.EOF = True Then
    GetARInvoiceDiscount = 0#
    Exit Function
  Else
    If IsNull(vntDate) Then vntDate = DateValue(Format(Now, "Short Date"))
    If IsDate(vntDate) Then
    Else
      vntDate = DateValue(Format(Now, "Short Date"))
    End If
    vntInvoiceDate = rsSales("AR SALE Date")
    PaymentTerms$ = NZ(rsSales("AR Sale Payment Terms"))
    'rsTerms.Index = "PrimaryKey"
    rsTerms.MoveFirst
    rsTerms.Find "[LIST PAY Description]='" & PaymentTerms$ & "'"
    If rsTerms.EOF = True Then
      GetARInvoiceDiscount = 0#
      Exit Function
    Else
      Discount# = rsTerms("LIST PAY Discount")
      If Discount# = 0 Then
        GetARInvoiceDiscount = 0#
        Exit Function
      Else
        Discount# = Discount# / 100
      End If
      DiscountDays# = rsTerms("LIST PAY Discount Days")
      Diff% = DateDiff("d", vntInvoiceDate, vntDate)
      If Diff% <= DiscountDays# Then
        DiscountAmount@ = rsSales("AR SALE Total") * Discount#
        GetARInvoiceDiscount = Round(CDbl(DiscountAmount@))
      Else
        GetARInvoiceDiscount = 0#
      End If
    End If
  End If

  rsSales.Close
  Set rsSales = Nothing
  rsTerms.Close
  Set rsTerms = Nothing
  db.Close
  Set db = Nothing
  
  Exit Function
GetARInvoiceDiscount_Error:
  Call ErrorLog("Sales Module", "GetARInvoiceDiscount", Now, Err.Number, Err.Description, True, db)
  Resume Next
  
  rsSales.Close
  Set rsSales = Nothing
  rsTerms.Close
  Set rsTerms = Nothing
  db.Close
  Set db = Nothing

End Function

Sub GetCustomerFinancials(CustomerID$)

  'Compute financial period balances for this customer

  'Get company's financial period information
  Dim Days&
  Dim AgeBy%
  Dim Period1%
  Dim Period2%
  Dim Period3%
  Dim Period4%
  Dim Balance#
  Dim TransDate As Variant
  Dim TransType$

  ''On Error GoTo CustomerFinancial_Error

  gCustomerPeriod1Balance# = 0
  gCustomerPeriod2Balance# = 0
  gCustomerPeriod3Balance# = 0
  gCustomerPeriod4Balance# = 0
  gCustomerTotalBalance# = 0

  Dim db As ADODB.Connection
  Set db = New ADODB.Connection
  db.CursorLocation = adUseClient
  db.Open gblADOProvider

  Dim rsCompany As ADODB.Recordset
  rsCompany.Open "SYS Company", db, adOpenStatic, adLockOptimistic, adCmdTable

  rsCompany.MoveFirst
  Period1% = rsCompany("SYS COM Sales Period 1")
  Period2% = rsCompany("SYS COM Sales Period 2")
  Period3% = rsCompany("SYS COM Sales Period 3")
  AgeBy% = IIf(IsNull(rsCompany("SYS COM Sales Age Invoices By")), 1, rsCompany("SYS COM Sales Age Invoices By"))
  '1 - Invoice Date  2 - Due Date

  'Go through AR Sales and get transactions for this customer
  '   with balances > 0
  
  Dim rsARPay As ADODB.Recordset
  rsARPay.Open "AR Payment Header", db, adOpenStatic, adLockOptimistic, adCmdTable
  
  Dim dn As ADODB.Recordset
  dn.Open "SELECT * FROM [AR Sales] where [AR SALE Customer ID] = '" & CustomerID$ & "' AND [AR SALE Balance Due] > 0 AND [AR SALE Posted YN] = TRUE", db, adOpenStatic, adLockOptimistic, adCmdText
  ''On Error Resume Next
  dn.MoveFirst
  If Err = 0 Then
    Do While Not dn.EOF
      'Get the balance
      Balance# = IIf(IsNull(dn("AR SALE Balance Due")), 0, dn("AR SALE Balance Due"))

      'Get Transaction Type to see if we sould increase or decrease the balance
      TransType$ = dn("AR SALE Document Type")
      Select Case TransType$
      Case "Invoice", "Sales Memo", "Beginning Balance", "Finance Charge"
      Case "Return", "Credit Memo"
        Balance# = 0
      Case Else
        Balance# = 0
      End Select

      'Get a date to age by
      If (AgeBy% = 1) Then 'Use Invoice Date
        TransDate = IIf(IsNull(dn("AR SALE Date")), Format("1/1/" & Format(Now, "yyyy"), "Short Date"), dn("AR SALE Date"))
      Else                 'Use Due Date
        TransDate = IIf(IsNull(dn("AR SALE Due Date")), Format("1/1/" & Format(Now, "yyyy"), "Short Date"), dn("AR SALE Due Date"))
      End If

      Days& = DateDiff("d", TransDate, Now)
      Select Case Days&
      Case Is < 0
        'Add it as current
        gCustomerPeriod1Balance# = gCustomerPeriod1Balance# + Balance#
      Case 0 To Period1%
        gCustomerPeriod1Balance# = gCustomerPeriod1Balance# + Balance#
      Case Period1% To Period2%
        gCustomerPeriod2Balance# = gCustomerPeriod2Balance# + Balance#
      Case Period2% To Period3%
        gCustomerPeriod3Balance# = gCustomerPeriod3Balance# + Balance#
      Case Else
        gCustomerPeriod4Balance# = gCustomerPeriod4Balance# + Balance#
      End Select
      dn.MoveNext
    Loop
  End If

  'Now do payments
  Dim dn2 As ADODB.Recordset
  dn2.Open "SELECT * FROM [AR PAYMENT Header] where [AR PAY Customer No] = '" & CustomerID$ & "' AND [AR PAY NSF] = False AND [AR PAY Posted YN] = TRUE", db, adOpenStatic, adLockOptimistic, adCmdText
  
  ''On Error Resume Next
  dn2.MoveFirst
  If Err = 0 Then
    Do While Not dn2.EOF
      'Get the balance
      Balance# = IIf(IsNull(dn2("AR PAY UnApplied Amount")), 0, dn2("AR PAY UnApplied Amount"))
      'Back out payment if type is return or NSF
      Select Case dn2("AR PAY Type")
      Case "NSF"
        Balance# = Balance# * -1
      End Select

      TransDate = IIf(IsNull(dn2("AR PAY Transaction Date")), Format("1/1/" & Format(Now, "yyyy"), "Short Date"), dn2("AR PAY Transaction Date"))
      
      Days& = DateDiff("d", TransDate, Now)

      Select Case Days&
      Case Is < 0
        'Don't use it
      Case 0 To Period1%
        gCustomerPeriod1Balance# = gCustomerPeriod1Balance# - Balance#
      Case Period1% To Period2%
        gCustomerPeriod2Balance# = gCustomerPeriod2Balance# - Balance#
      Case Period2% To Period3%
        gCustomerPeriod3Balance# = gCustomerPeriod3Balance# - Balance#
      Case Else
        gCustomerPeriod4Balance# = gCustomerPeriod4Balance# - Balance#
      End Select
      dn2.MoveNext
    Loop
  End If

  gCustomerTotalBalance# = gCustomerPeriod1Balance# + gCustomerPeriod2Balance# + gCustomerPeriod3Balance# + gCustomerPeriod4Balance#

    
CustomerFinancials_Exit:
  dn.Close
  Set dn = Nothing
  dn2.Close
  Set dn2 = Nothing
  rsCompany.Close
  Set rsCompay = Nothing
  rsARPay.Close
  Set rsARPay = Nothing
  db.Close
  Set db = Nothing
  Exit Sub

CustomerFinancial_Error:
  Call ErrorLog("Sales Module", "GetCustomerFinancials", Now, Err.Number, Err.Description, True, db)
  dn.Close
  Set dn = Nothing
  dn2.Close
  Set dn2 = Nothing
  rsCompany.Close
  Set rsCompay = Nothing
  rsARPay.Close
  Set rsARPay = Nothing
  db.Close
  Set db = Nothing
  Exit Sub

  dn.Close
  Set dn = Nothing
  dn2.Close
  Set dn2 = Nothing
  rsCompany.Close
  Set rsCompay = Nothing
  rsARPay.Close
  Set rsARPay = Nothing
  db.Close
  Set db = Nothing


End Sub

Sub MonthEndSales(db As ADODB.Connection)
On Error GoTo MonthEndSales_Error
  'Check all invoices that are not cleared yet
  'If Balance = 0 and all payments with an amount
  '  applied to this invoice are fully applied
  'then
  '  clear this invoice
  'end if

  'Dim db As ADODB.Connection
  'Set db = New ADODB.Connection
  'db.CursorLocation = adUseClient
  'db.Open gblADOProvider
  
  'Dim rs As ADODB.Recordset
  'Dim rs2 As ADODB.Recordset
  
  ''On Error GoTo MonthEndSales_Error
  'Set rs = New ADODB.Recordset
  'rs.Open "Select * from [AR Sales] where [AR SALE Cleared] = 0 and [AR SALE Balance Due] < .01", db, adOpenStatic, adLockOptimistic, adCmdText
  db.BeginTrans
  db.Execute "UPDATE [AR Sales] SET [AR SALE Cleared]=True WHERE [AR SALE Cleared] = 0 " & _
  "AND [AR SALE Balance Due] < .01", , adCmdText
  ''On Error Resume Next
  'rs.MoveFirst
  'If Not rs.EOF Then
  '  Do While Not rs.EOF
  '    rs("AR SALE Cleared") = True
  '    rs.Update
  '    rs.MoveNext
  '  Loop
  'End If
  'Set rs2 = New ADODB.Recordset
  db.Execute "UPDATE [AR PAYMENT Header] SET [AR PAY Cleared]=True WHERE " & _
  "[AR PAY Cleared] = 0 and [AR PAY Unapplied Amount] < .01", , adCmdText
  'rs2.Open "SELECT * FROM [AR PAYMENT Header] where [AR PAY Cleared] = 0 and [AR PAY Unapplied Amount] < .01", db, adOpenStatic, adLockOptimistic, adCmdText
  
  ''On Error Resume Next
  'rs2.MoveFirst
  'If Not rs2.EOF Then
  '  Do While Not rs.EOF
  '    rs2("AR PAY Cleared") = True
  '    rs2.Update
  '    rs2.MoveNext
  '  Loop
  'End If

  'rs.Close
  'Set rs = Nothing
  'rs2.Close
  'Set rs2 = Nothing
  'db.Close
  'Set db = Nothing
  db.CommitTrans
  Exit Sub

MonthEndSales_Error:
  Call ErrorLog("Sales Module", "MonthEndSales", Now, Err.Number, Err.Description, True, db)
  db.RollbackTrans
  'rs.Close
  'Set rs = Nothing
  'rs2.Close
  'Set rs2 = Nothing
  'db.Close
  'Set db = Nothing

  'Exit Sub

  'rs.Close
  'Set rs = Nothing
  'rs2.Close
  'Set rs2 = Nothing
  'db.Close
  'Set db = Nothing


End Sub

Function PostCreditMemo(DocumentKey&, intShowError As Integer, Optional db As ADODB.Connection) As Integer
Dim DBload As Boolean

  ''On Error GoTo PostCreditMemo_Error

  Dim msg$
  Dim title$

DBload = False
If db Is Nothing Then
  Set db = New ADODB.Connection
  db.CursorLocation = adUseServer
  db.Open gblADOProvider
  DBload = True
End If
  
  Dim rsCompany As ADODB.Recordset
  Set rsCompany = New ADODB.Recordset
  rsCompany.Open "SELECT [SYS COM GL Post By Date],[SYS COM Sales Acct Default]," & _
  "[SYS COM Inventory Cost Method Last YN],[SYS COM Sales AR Acct],[SYS COM Sales Discount Acct]," & _
  "[SYS COM Sales Sales Acct],[SYS COM Sales Misc Acct],[SYS COM Sales Sales Tax]," & _
  "[SYS COM Sales Freight Acct],[SYS COM Sales COGS Acct],[SYS COM Sales Inventory Acct] " & _
  "FROM [SYS Company]", db, adOpenKeyset, adLockOptimistic, adCmdText
  
  'Dim rsGLWorkDetail As ADODB.Recordset
  'Set rsGLWorkDetail = New ADODB.Recordset
  'rsGLWorkDetail.Open "[GL Work Detail]", db, adOpenStatic, adLockOptimistic, adCmdTable

  Dim rsSales As ADODB.Recordset
  Set rsSales = New ADODB.Recordset
  'rsSales.Open "[AR Sales]", db, adOpenStatic, adLockOptimistic, adCmdTable
  rsSales.Open "SELECT [AR SALE Date],[AR SALE Customer ID],[AR SALE Total],[AR SALE Ext Document #]," & _
  "[AR SALE Check Acct ID],[AR SALE Billing Customer],[AR SALE Bill To],[AR SALE Description]," & _
  "[AR SALE Discount Amount],[AR SALE Sales Tax],[AR SALE Tax Group],[AR SALE Taxable Subtotal]," & _
  "[AR SALE Freight],[AR SALE Document #] FROM [AR Sales] WHERE [AR SALE Document #]=" & DocumentKey&, db, adOpenKeyset, adLockOptimistic, adCmdText
  
  'rsSales.Index = "PrimaryKey"
  'rsSales.MoveFirst
  'rsSales.Find "[AR SALE Document #]=" & DocumentKey&

  'Post by 1 - system date or 2 - Transaction date?
  Dim PostDate%
  PostDate% = rsCompany("SYS COM GL Post By Date")

  'Set Invoice Type
  Dim InvoiceType$
  InvoiceType$ = "Credit Memo"

  'Set Post Date
  Dim TranDate As Variant
  If PostDate% = 1 Then
    TranDate = DateValue(Format(Now, "Short Date"))
  Else
    TranDate = DateValue(rsSales("AR SALE Date"))
  End If
  
  'Verify period can be posted to Send TranDate Return PeriodToPost and PeriodClosed
  Dim PeriodToPost%
  Dim PeriodClosed%
  Call VerifyPeriod(TranDate, PeriodToPost%, PeriodClosed%, db)
  
  'Is period open?
  If PeriodClosed% = True Then
    MsgBox "Unable to post transaction to a closed period.", , "Post Invoice Error"
    GoTo UnableToPostCreditHere
  End If

  ''On Error GoTo PostCreditMemo_Error
  
  ' clear any GL Work records
  'Dim cmdtemp As ADODB.Recordset
  'Set cmdtemp = New ADODB.Recordset
  'cmdtemp.Open "DELETE DISTINCTROW * FROM [GL Work Detail]", db, , , adCmdText
  'cmdtemp.Close
  'Set cmdtemp = Nothing
  db.Execute "DELETE DISTINCTROW * FROM [GL Work Detail]", , adCmdText


  ' update customer stats
  'Dim rsCustomer As ADODB.Recordset
  'Set rsCustomer = New ADODB.Recordset
  'rsCustomer.Open "SELECT * FROM [AR Customer] WHERE [AR CUST Customer ID] = '" & rsSales("AR SALE Customer ID") & "'", db, adOpenStatic, adLockOptimistic, adCmdText

  ' if we are to post by customer get the Customer Sales Acct
  'Dim SalesAcctDefault%
  'Dim CustomerSalesAcct$
  '1-Post using customer accounts; 2-Post using item accounts

  'SalesAcctDefault% = rsCompany("SYS COM Sales Acct Default")
  'CustomerSalesAcct$ = ""
  'If SalesAcctDefault% = 1 Then
  '  CustomerSalesAcct$ = IIf(IsNull(rsCustomer("AR CUST Sales Account")), "", rsCustomer("AR CUST Sales Account"))
  'End If

  'rsCustomer.Update
 
  '--------------------------------------------------
  ' New AR Cross Payment and AR Payment Header
    'Dim rsARPaymentHeader As ADODB.Recordset
    'Set rsARPaymentHeader = New ADODB.Recordset
    'rsARPaymentHeader.Open "[AR Payment Header]", db, adOpenStatic, adLockOptimistic, adCmdTable
  
  'Dim rsARCross As ADODB.Recordset
  'Set rsARCross = New ADODB.Recordset
  'rsARCross.Open "[AR Payment Invoice Cross Reference]", db, adOpenStatic, adLockOptimistic, adCmdTable

  Dim PaymentID&

  If rsSales("AR SALE Total") > 0 Then
    
    '  write Payment Header
    'rsARPaymentHeader.AddNew
    '  rsARPaymentHeader("AR PAY Type") = "Credit Memo"
    '  rsARPaymentHeader("AR PAY Check No") = "CM " & rsSales("AR SALE Ext Document #") & ""
    '  rsARPaymentHeader("AR PAY Customer No") = rsSales("AR SALE Customer ID") & ""
    '  rsARPaymentHeader("AR PAY Transaction Date") = rsSales("AR SALE Date")
    '  rsARPaymentHeader("AR PAY Amount") = rsSales("AR SALE Total")
    '  rsARPaymentHeader("AR PAY UnApplied Amount") = rsSales("AR SALE Total")  'Cannot have unapplied amounts here
    '  ----rsARPaymentHeader("AR PAY Bank Account") = rsSales("AR SALE Check Acct ID")
    '  rsARPaymentHeader("AR PAY Status") = "Posted"
    '  rsARPaymentHeader("AR PAY NSF") = False
    '  rsARPaymentHeader("AR PAY Posted YN") = True
    '  rsARPaymentHeader("AR PAY Cleared") = False
    'rsARPaymentHeader.Update
    ' end of write payment header
    '  PaymentID& = rsARPaymentHeader("AR PAY ID")
      
      SQLstatement = "INSERT INTO [AR Payment Header]"
      SQLstatement = SQLstatement & " ([AR PAY Type],[AR PAY Check No],[AR PAY Customer No],[AR PAY Transaction Date],[AR PAY Amount],[AR PAY UnApplied Amount],[AR PAY Status],[AR PAY Posted YN],[AR PAY NSF],[AR PAY Cleared])"
      SQLstatement = SQLstatement & " VALUES ('Credit Memo','CM " & rsSales("AR SALE Ext Document #") & "','" & CStr(rsSales("AR SALE Customer ID")) & "',#" & rsSales("AR SALE Date") & "#," & rsSales("AR SALE Total") & "," & rsSales("AR SALE Total") & ",'Posted',True,False,False)"
      db.Execute SQLstatement
    
      Dim rsARPaymentHeader As ADODB.Recordset
      Set rsARPaymentHeader = New ADODB.Recordset
      rsARPaymentHeader.Open "SELECT [AR PAY ID] FROM [AR Payment Header] WHERE [AR PAY Check No]='CM " & rsSales("AR SALE Ext Document #") & "'", db, adOpenKeyset, adLockReadOnly, adCmdText
      If rsARPaymentHeader.RecordCount > 1 Then
        rsARPaymentHeader.MoveLast
        PaymentID& = rsARPaymentHeader("AR PAY ID")
      Else
        PaymentID& = rsARPaymentHeader("AR PAY ID")
      End If
      rsARPaymentHeader.Close
      Set rsARPaymentHeader = Nothing

  End If
  ' end of write AR Cross reference Record


  ' End of AR Cross Payment and AR Payment Header
  '--------------------------------------------------


  'Dim rsGLTrans As ADODB.Recordset
  '  Set rsGLTrans = New ADODB.Recordset
  'rsGLTrans.Open "[GL Transaction]", db, adOpenStatic, adLockOptimistic, adCmdTable

  ' write GL Transaction Header
  Dim refr$
  Dim desc$
  Dim NewNumber&
  'rsGLTrans.AddNew
      SQLstatement = "INSERT INTO [GL Transaction]"
      SQLstatement = SQLstatement & " ([GL TRANS Document #],[GL TRANS Type],[GL TRANS Date],"
      SQLstatement = SQLstatement & " [GL TRANS Reference],[GL TRANS Amount],[GL TRANS Posted YN],"
      SQLstatement = SQLstatement & " [GL TRANS Description],[GL TRANS Source],[GL TRANS System Generated])"
    
  '  rsGLTrans("GL TRANS Document #") = "Credit Memo " & rsSales("AR SALE Ext Document #")
  '  rsGLTrans("GL TRANS Type") = "Credit Memo"
    Dim TempStr As String
    
    ' gl post date
    If PostDate% = 1 Then
      TempStr = Format(Now, "Short Date")
    Else
      TempStr = rsSales("AR SALE Date")
    End If
    
    SQLstatement = SQLstatement & " VALUES ('Credit Memo " & rsSales("AR SALE Ext Document #") & "','Credit Memo',#" & TempStr & "#,"
    
    If IsNull(rsSales("AR SALE Billing Customer")) Then
        refr$ = rsSales("AR SALE Bill To") & ""
    Else
        refr$ = rsSales("AR SALE Billing Customer") & ""
    End If
    
    desc$ = IIf(IsNull(rsSales("AR SALE Description")), "", rsSales("AR SALE Description"))
    If Len(Trim$(desc$)) = 0 Then
      desc$ = "Credit Memo " & rsSales("AR SALE Ext Document #")
    End If
    
      SQLstatement = SQLstatement & "'" & refr$ & "'," & rsSales("AR SALE Total") & ",1,"
      SQLstatement = SQLstatement & "'" & desc$ & "','Credit Memo " & rsSales("AR SALE Ext Document #") & "',True)"
      'Debug.Print SQLstatement
      
      db.Execute SQLstatement
    
    'rsGLTrans("GL TRANS Reference") = refr$
    'rsGLTrans("GL TRANS Amount") = rsSales("AR SALE Total")
    'rsGLTrans("GL TRANS Posted YN") = 1
    'rsGLTrans("GL TRANS Description") = desc$
    'rsGLTrans("GL TRANS Source") = "Credit Memo " & rsSales("AR SALE Ext Document #")
    'rsGLTrans("GL TRANS System Generated") = True
  'rsGLTrans.Update
    'NewNumber& = rsGLTrans("GL TRANS Number")
  Dim rsGLTrans As ADODB.Recordset
  Set rsGLTrans = New ADODB.Recordset
  rsGLTrans.Open "SELECT [GL TRANS Number] FROM [GL Transaction]WHERE [GL TRANS Document #]='" & "Credit Memo " & rsSales("AR SALE Ext Document #") & "'", db, adOpenKeyset, adLockReadOnly, adCmdText
      NewNumber& = rsGLTrans("GL TRANS Number")
  rsGLTrans.Close
  Set rsGLTrans = Nothing
  ' write GL Transaction Detail


  ' update GL
  '  Transaction #1
  '-----------------------------------------------------------------------
  ' Invoice with Payment GL Affected Accounts
  '
  '                  Debit   Credit   Source
  '                  -----   ------   ------
  ' AR                         X      Pref - Sales
  ' CASH                       X      Bank - Cash Acct
  ' Discount                   X      Pref - Sales
  ' Sales              X              Item - Inventory or Customer - gSalesAcctDefault%
  ' Sales Tax          X              Sys Tax
  ' Freight Income     X              Pref - Sales
  ' The following entry is only valid if the payment exceeds the invoice total
  ' AR                 X              Pref - Sales
  '
  '-----------------------------------------------------------------------

  ' Credits
  ' AR

  If rsSales("AR SALE Total") > 0 Then
      SQLstatement = "INSERT INTO [GL Work Detail]"
      SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
      SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & rsCompany("SYS COM Sales AR Acct") & "" & "',0," & rsSales("AR SALE Total") & ")"
      db.Execute SQLstatement
    'rsGLWorkDetail.AddNew
    '  rsGLWorkDetail("GW TRANSD Number") = NewNumber&
    '  rsGLWorkDetail("GW TRANSD Account") = rsCompany("SYS COM Sales AR Acct") ----
    '  rsGLWorkDetail("GW TRANSD Debit Amount") = 0
    '  rsGLWorkDetail("GW TRANSD Credit Amount") = rsSales("AR SALE Total")
    '  rsGLWorkDetail("GW TRANSD Project") = ""
    'rsGLWorkDetail.Update
  End If
  
  ' Discount Amount
  If rsSales("AR SALE Discount Amount") > 0 Then
      SQLstatement = "INSERT INTO [GL Work Detail]"
      SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
      SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & rsCompany("SYS COM Sales Discount Acct") & "" & "',0," & rsSales("AR SALE Discount Amount") & ")"
      db.Execute SQLstatement
    'rsGLWorkDetail.AddNew
    '  rsGLWorkDetail("GW TRANSD Number") = NewNumber&
    '  rsGLWorkDetail("GW TRANSD Account") = rsCompany("SYS COM Sales Discount Acct")
    '  rsGLWorkDetail("GW TRANSD Debit Amount") = 0
    '  rsGLWorkDetail("GW TRANSD Credit Amount") = rsSales("AR SALE Discount Amount")
    '  rsGLWorkDetail("GW TRANSD Project") = ""
    'rsGLWorkDetail.Update
  End If
  

  'Debits

  ' Sales Tax
  Dim GLTaxAcct$
  Dim TaxGroup$
  Dim TaxPercent#
  Dim TaxID$
  Dim Taxtotal#
  
  If rsSales("AR SALE Sales Tax") > 0 Then
    GLTaxAcct$ = ""
    TaxGroup$ = IIf(IsNull(rsSales("AR SALE Tax Group")), rsCompany("SYS COM Sales Sales Tax"), rsSales("AR SALE Tax Group"))
    Dim rsTaxGroupDetail As ADODB.Recordset
    
    Set rsTaxGroupDetail = New ADODB.Recordset
    rsTaxGroupDetail.Open "SELECT * FROM [SYS Tax Group Detail] where [SYS TAXGRPD Group ID] = '" & TaxGroup$ & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    
    If rsTaxGroupDetail.RecordCount = 0 Then
    Else
      rsTaxGroupDetail.MoveFirst
      Do While rsTaxGroupDetail.EOF = False
        TaxPercent# = 0
        TaxID$ = rsTaxGroupDetail("SYS TAXGRPD Tax ID")
        TaxPercent# = LookRecord("[SYS Tax Percent]", "[SYS Tax]", db, "[SYS Tax ID] = '" & TaxID$ & "'")
        GLTaxAcct$ = LookRecord("[SYS Tax Account]", "[SYS Tax]", db, "[SYS Tax ID] = '" & TaxID$ & "'")
        
        If GLTaxAcct$ = "" Then
          GLTaxAcct$ = rsCompany("SYS COM Sales Sales Acct")
        End If
        Taxtotal# = Round(rsSales("AR SALE Taxable Subtotal") * (TaxPercent# / 100))
        SQLstatement = "INSERT INTO [GL Work Detail]"
        SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
        SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & GLTaxAcct$ & "'," & Taxtotal# & ",0)"
        db.Execute SQLstatement
        
        'rsGLWorkDetail.AddNew
        '  rsGLWorkDetail("GW TRANSD Number") = NewNumber&
        '  rsGLWorkDetail("GW TRANSD Account") = GLTaxAcct$
        '  rsGLWorkDetail("GW TRANSD Debit Amount") = Round(rsSales("AR SALE Taxable Subtotal") * (TaxPercent# / 100))
        '  rsGLWorkDetail("GW TRANSD Credit Amount") = 0
        '  rsGLWorkDetail("GW TRANSD Project") = ""
        'rsGLWorkDetail.Update

        rsTaxGroupDetail.MoveNext
      Loop
      rsTaxGroupDetail.Close
      Set rsTaxGroupDetail = Nothing
    End If
  End If
  ' Sales Tax


  ' Freight Income
  If rsSales("AR SALE Freight") > 0 Then
    SQLstatement = "INSERT INTO [GL Work Detail]"
    SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
    SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & rsCompany("SYS COM Sales Freight Acct") & "" & "'," & rsSales("AR SALE Freight") & ",0)"
    db.Execute SQLstatement
    'rsGLWorkDetail.AddNew
    '  rsGLWorkDetail("GW TRANSD Number") = NewNumber&
    '  rsGLWorkDetail("GW TRANSD Account") = rsCompany("SYS COM Sales Freight Acct")
    '  rsGLWorkDetail("GW TRANSD Debit Amount") = rsSales("AR SALE Freight")
    '  rsGLWorkDetail("GW TRANSD Credit Amount") = 0
    '  rsGLWorkDetail("GW TRANSD Project") = ""
    'rsGLWorkDetail.Update
  End If

  ' Sales Increase
  
  Dim Longer&
  Dim InventoryAcct$
  
  ' item
  Dim rsDetail As ADODB.Recordset
  Set rsDetail = New ADODB.Recordset

  Longer& = 0
  rsDetail.Open "SELECT [AR SALED Posting Account],[AR SALED Item Total] FROM [AR Sales Detail] where [AR SALED Document #] = " & rsSales("AR SALE Document #"), db, adOpenStatic, adLockOptimistic, adCmdText
  
  If rsDetail.RecordCount = 0 Then
    ' no detail data
  Else
  rsDetail.MoveFirst
    ' may have detail data
    Do While rsDetail.EOF = False
    SQLstatement = "INSERT INTO [GL Work Detail]"
    SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
    SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & rsDetail("AR SALED Posting Account") & "" & "'," & rsDetail("AR SALED Item Total") & ",0)"
    db.Execute SQLstatement
      'rsGLWorkDetail.AddNew
      '  rsGLWorkDetail("GW TRANSD Number") = NewNumber&
      '  rsGLWorkDetail("GW TRANSD Account") = rsDetail("AR SALED Posting Account")
      '  rsGLWorkDetail("GW TRANSD Debit Amount") = rsDetail("AR SALED Item Total")
      '  rsGLWorkDetail("GW TRANSD Credit Amount") = 0
      '  rsGLWorkDetail("GW TRANSD Project") = ""
      'rsGLWorkDetail.Update

      rsDetail.MoveNext
      'If Err = 3021 Then Exit Do
    Loop
   rsDetail.Close
   Set rsDetail = Nothing
  End If
  
  ' post GL entry for this Credit Memo
  
  Dim Success%
  Success% = PostGLWorkDetail(TranDate, NewNumber&, db)
  If Success% = False Then
    MsgBox "An error occurred writing GL Transaction!", , "Error"
    PostCreditMemo = False
    Exit Function
  End If

PostCreditMemo_Exit:

  PostCreditMemo = True
  
  'rsGLWorkDetail.Close
  'Set rsGLWorkDetail = Nothing
  'rsDetail.Close
  'Set rsDetail = Nothing
  'rsTaxGroupDetail.Close
  'Set rsTaxGroupDetail = Nothing
  'rsGLTrans.Close
  'Set rsGLTrans = Nothing
  rsCompany.Close
  Set rsCompany = Nothing
  rsSales.Close
  Set rsSales = Nothing
  'rsARPaymentHeader.Close
  'Set rsARPaymentHeader = Nothing
  'rsARCross.Close
  'Set rsARCross = Nothing
  'rsCustomer.Close
  'Set rsCustomer = Nothing
  'rsCompany.Close
  'Set rsCompany = Nothing
If DBload = True Then
  db.Close
  Set db = Nothing
End If
  Exit Function

PostCreditMemo_Error:
  Call ErrorLog("Sales Module", "PostCreditMemo", Now, Err.Number, Err.Description, intShowError, db)
  PostCreditMemo = False
  'rsGLWorkDetail.Close
  'Set rsGLWorkDetail = Nothing
  'rsDetail.Close
  'Set rsDetail = Nothing
  'rsTaxGroupDetail.Close
  'Set rsTaxGroupDetail = Nothing
  'rsGLTrans.Close
  'Set rsGLTrans = Nothing
  rsCompany.Close
  Set rsCompany = Nothing
  rsSales.Close
  Set rsSales = Nothing
  'rsARPaymentHeader.Close
  'Set rsARPaymentHeader = Nothing
  'rsARCross.Close
  'Set rsARCross = Nothing
  'rsCustomer.Close
  'Set rsCustomer = Nothing
  'rsCompany.Close
  'Set rsCompany = Nothing
If DBload = True Then
  db.Close
  Set db = Nothing
End If
  Exit Function
  'Resume Next

UnableToPostCreditHere:
  PostCreditMemo = False
  'rsGLWorkDetail.Close
  'Set rsGLWorkDetail = Nothing
  'rsDetail.Close
  'Set rsDetail = Nothing
  'rsTaxGroupDetail.Close
  'Set rsTaxGroupDetail = Nothing
  'rsGLTrans.Close
  'Set rsGLTrans = Nothing
  rsCompany.Close
  Set rsCompany = Nothing
  rsSales.Close
  Set rsSales = Nothing
  'rsARPaymentHeader.Close
  'Set rsARPaymentHeader = Nothing
  'rsARCross.Close
  'Set rsARCross = Nothing
  'rsCustomer.Close
  'Set rsCustomer = Nothing
  'rsCompany.Close
  'Set rsCompany = Nothing
If DBload = True Then
  db.Close
  Set db = Nothing
End If
  Exit Function

  'rsGLWorkDetail.Close
  'Set rsGLWorkDetail = Nothing
  'rsDetail.Close
  'Set rsDetail = Nothing
  'rsTaxGroupDetail.Close
  'Set rsTaxGroupDetail = Nothing
  'rsGLTrans.Close
  'Set rsGLTrans = Nothing
  rsCompany.Close
  Set rsCompany = Nothing
  rsSales.Close
  Set rsSales = Nothing
  'rsARPaymentHeader.Close
  'Set rsARPaymentHeader = Nothing
  'rsARCross.Close
  'Set rsARCross = Nothing
  'rsCustomer.Close
  'Set rsCustomer = Nothing
  'rsCompany.Close
  'Set rsCompany = Nothing
If DBload = True Then
  db.Close
  Set db = Nothing
End If

End Function

Function PostInvoice(DocumentKey&, intShowError As Integer, Optional db As ADODB.Connection) As Integer
Dim DBload As Boolean
  '''On Error GoTo PostInvoice_Error

  Dim msg$
  Dim title$
  
DBload = False
If db Is Nothing Then
  Set db = New ADODB.Connection
  db.CursorLocation = adUseServer
  db.Open gblADOProvider
  DBload = True
End If

  Dim rsCompany As ADODB.Recordset
  Set rsCompany = New ADODB.Recordset
  rsCompany.Open "SELECT [SYS COM GL Post By Date],[SYS COM Sales Acct Default]," & _
  "[SYS COM Inventory Cost Method Last YN],[SYS COM Sales AR Acct],[SYS COM Sales Discount Acct]," & _
  "[SYS COM Sales Sales Acct],[SYS COM Sales Misc Acct],[SYS COM Sales Sales Tax]," & _
  "[SYS COM Sales Freight Acct],[SYS COM Sales COGS Acct],[SYS COM Sales Inventory Acct] " & _
  "FROM [SYS Company]", db, adOpenStatic, adLockOptimistic, adCmdText
  
  rsCompany.MoveFirst
  
  'Dim rsGLWorkDetail As ADODB.Recordset
  'Set rsGLWorkDetail = New ADODB.Recordset
  'rsGLWorkDetail.Open "[GL Work Detail]", db, adOpenStatic, adLockOptimistic, adCmdTable
  
  Dim rsSales As ADODB.Recordset
  Set rsSales = New ADODB.Recordset
  rsSales.Open "SELECT * FROM [AR Sales] WHERE [AR SALE Document #]=" & DocumentKey&, db, adOpenStatic, adLockOptimistic, adCmdText
  
  'rsSales.Index = "PrimaryKey"
  'rsSales.MoveFirst
  'rsSales.Find "[AR SALE Document #]=" & DocumentKey&

  'Post by 1 - system date or 2 - Transaction date?
  Dim PostDate%
  PostDate% = rsCompany("SYS COM GL Post By Date")

  'Set Invoice Type
  Dim InvoiceType$
  InvoiceType$ = rsSales("AR SALE Document Type")

  'Set Post Date
  Dim TranDate As Variant
  If PostDate% = 1 Then
    TranDate = DateValue(Format(Now, "Short Date"))
  Else
    TranDate = DateValue(rsSales("AR SALE Date"))
  End If
  
  'Verify period can be posted to
  'Send TranDate
  'Return PeriodToPost and PeriodClosed
  Dim PeriodToPost%
  Dim PeriodClosed%
  Call VerifyPeriod(TranDate, PeriodToPost%, PeriodClosed%, db)
  
  'Is period open?
  If PeriodClosed% = True Then
    MsgBox "Unable to post transaction to a closed period.", , "Post Invoice Error"
    GoTo UnableToPostHere
  End If

  'On Error GoTo PostInvoice_Error
  
  ' clear any GL Work records
  'Dim cmdtemp As ADODB.Recordset
  'Set cmdtemp = New ADODB.Recordset
  'cmdtemp.Open "DELETE DISTINCTROW * FROM [GL Work Detail]", db, , , adCmdText
  db.Execute "DELETE DISTINCTROW * FROM [GL Work Detail]", , adCmdText
  'cmdtemp.Close
  'Set cmdtemp = Nothing


  Dim BankKey$
  BankKey$ = IIf(IsNull(rsSales("AR SALE Check Acct ID")), "", rsSales("AR SALE Check Acct ID"))

  Dim AppliedAmount@
  If IsNull(rsSales("AR SALE Amount Paid")) Then
    AppliedAmount@ = 0
  Else
    AppliedAmount@ = rsSales("AR SALE Amount Paid")
  End If

  ' update customer stats
  Dim rsCustomer As ADODB.Recordset
  Set rsCustomer = New ADODB.Recordset
  rsCustomer.Open "SELECT * FROM [AR Customer] WHERE [AR CUST Customer ID] = '" & rsSales("AR SALE Customer ID") & "'", db, adOpenStatic, adLockOptimistic, adCmdText

  ' if we are to post by customer get the Customer Sales Acct
  Dim SalesAcctDefault%
  Dim CustomerSalesAcct$
  '1-Post using customer accounts; 2-Post using item accounts

  Dim CurrentBalance@

  SalesAcctDefault% = rsCompany("SYS COM Sales Acct Default")
  CustomerSalesAcct$ = ""
  If SalesAcctDefault% = 1 Then
    CustomerSalesAcct$ = IIf(IsNull(rsCustomer("AR CUST Sales Account")), "", rsCustomer("AR CUST Sales Account"))
  End If

    rsCustomer("AR CUST Payments YTD") = rsCustomer("AR CUST Payments YTD") + rsSales("AR SALE Amount Paid")
    rsCustomer("AR CUST Payments Lifetime") = rsCustomer("AR CUST Payments Lifetime") + rsSales("AR SALE Amount Paid")
    rsCustomer("AR CUST Sales YTD") = rsCustomer("AR CUST Sales YTD") + rsSales("AR SALE Total")
    rsCustomer("AR CUST Sales Lifetime") = rsCustomer("AR CUST Sales Lifetime") + rsSales("AR SALE Total")
    rsCustomer("AR CUST Invoices Lifetime") = rsCustomer("AR CUST Invoices Lifetime") + 1
    rsCustomer("AR CUST Invoices YTD") = rsCustomer("AR CUST Invoices YTD") + 1

    ' Update current Balance - if not Paid in full
    If rsSales("AR SALE Total") - rsSales("AR SALE Amount Paid") > 0 Then
      CurrentBalance@ = IIf(IsNull(rsCustomer("AR CUST Financial Period 1")), 0, rsCustomer("AR CUST Financial Period 1"))
      CurrentBalance@ = CurrentBalance@ + (rsSales("AR SALE Total") - rsSales("AR SALE Amount Paid"))
      rsCustomer("AR CUST Financial Period 1") = CurrentBalance@
      If CurrentBalance@ > IIf(IsNull(rsCustomer("AR CUST Highest Balance")), 0, rsCustomer("AR CUST Highest Balance")) Then rsCustomer("AR CUST Highest Balance") = CurrentBalance@
    End If
  rsCustomer.Update
 
  '--------------------------------------------------

  Dim PaymentID&

  If AppliedAmount@ > 0 Then
    ' New AR Cross Payment and AR Payment Header
          
    'rsARPaymentHeader.AddNew
    '  rsARPaymentHeader("AR PAY Type") = "Payment"
    '  rsARPaymentHeader("AR PAY Check No") = CStr(rsSales("AR SALE Check Number"))
    '  rsARPaymentHeader("AR PAY Customer No") = CStr(rsSales("AR SALE Customer ID"))
    '  rsARPaymentHeader("AR PAY Transaction Date") = rsSales("AR SALE Date")
    '  rsARPaymentHeader("AR PAY Amount") = rsSales("AR SALE Amount Paid")
    '  rsARPaymentHeader("AR PAY UnApplied Amount") = 0  'Cannot have unapplied amounts here
    '  rsARPaymentHeader("AR PAY Bank Account") = BankKey$
    '  rsARPaymentHeader("AR PAY Status") = "Posted"
    '  rsARPaymentHeader("AR PAY Posted YN") = True
    '  rsARPaymentHeader("AR PAY NSF") = False
    '  rsARPaymentHeader("AR PAY Cleared") = False
    'this request might returned a wrong value in multi-tier development
    'rsARPaymentHeader.Update
      
      SQLstatement = "INSERT INTO [AR Payment Header]"
      SQLstatement = SQLstatement & " ([AR PAY Type],[AR PAY Check No],[AR PAY Customer No],[AR PAY Transaction Date],[AR PAY Amount],[AR PAY UnApplied Amount],[AR PAY Bank Account],[AR PAY Status],[AR PAY Posted YN],[AR PAY NSF],[AR PAY Cleared])"
      SQLstatement = SQLstatement & " VALUES ('Payment','" & CStr(rsSales("AR SALE Check Number")) & "','" & CStr(rsSales("AR SALE Customer ID")) & "',#" & rsSales("AR SALE Date") & "#," & rsSales("AR SALE Amount Paid") & ",0," & BankKey$ & ",'Posted',True,False,False)"
      db.Execute SQLstatement
                                  
    Dim rsARPaymentHeader As ADODB.Recordset
    Set rsARPaymentHeader = New ADODB.Recordset
    rsARPaymentHeader.Open "SELECT [AR PAY ID] FROM [AR Payment Header] WHERE [AR PAY Check No]='" & CStr(rsSales("AR SALE Check Number")) & "'", db, adOpenKeyset, adLockReadOnly, adCmdText
    If rsARPaymentHeader.RecordCount > 1 Then
      rsARPaymentHeader.MoveLast
      PaymentID& = rsARPaymentHeader("AR PAY ID")
    Else
      PaymentID& = rsARPaymentHeader("AR PAY ID")
    End If
    rsARPaymentHeader.Close
    Set rsARPaymentHeader = Nothing
    ' end of write payment header

    ' write AR Cross reference Record
    'Dim rsARCross As ADODB.Recordset
    'Set rsARCross = New ADODB.Recordset
    'rsARCross.Open "[AR Payment Invoice Cross Reference]", db, adOpenStatic, adLockOptimistic, adCmdTable
    'rsARCross.AddNew
    '  rsARCross("AR CROSS Payment ID") = PaymentID&
    '  rsARCross("AR CROSS Payed ID") = rsSales("AR SALE Document #")
    '  rsARCross("AR CROSS Discount Taken") = 0
    '  rsARCross("AR CROSS Write Off Amount") = 0
    '  rsARCross("AR CROSS Applied Amount") = AppliedAmount@
    '  rsARCross("AR CROSS Cleared") = False
    'rsARCross.Update
    
      SQLstatement = "INSERT INTO [AR Payment Invoice Cross Reference]"
      SQLstatement = SQLstatement & " ([AR CROSS Payment ID],[AR CROSS Payed ID],[AR CROSS Discount Taken],[AR CROSS Write Off Amount],[AR CROSS Applied Amount],[AR CROSS Cleared])"
      SQLstatement = SQLstatement & " VALUES (" & PaymentID& & "," & rsSales("AR SALE Document #") & ",0,0," & AppliedAmount@ & ",False)"
      db.Execute SQLstatement
                                  
  End If
  ' end of write AR Cross reference Record


  ' End of New AR Cross Payment and AR Payment Header
  '--------------------------------------------------

  'Inventory Updates
  Dim rsInventory As ADODB.Recordset
  Dim rsDetail As ADODB.Recordset
    
  Set rsDetail = New ADODB.Recordset
  rsDetail.Open "SELECT [AR SALED Item ID],[AR SALED Item Cost],[AR SALED Qty],[AR SALED Item Total],[AR SALED Units],[AR SALED Posting Account],[AR SALED Row Type] FROM [AR Sales Detail] where [AR SALED Document #] = " & rsSales("AR SALE Document #"), db, adOpenKeyset, adLockOptimistic, adCmdText

  Dim QtySold#

  Dim CostMethod%
  '-1 - Use Last Cost Method;  0 - Use Average Costing
  CostMethod% = rsCompany("SYS COM Inventory Cost Method Last YN")
  
  If rsDetail.RecordCount = 0 Then
    ' no detail data
  Else
    rsDetail.MoveFirst
    ' may have detail data
    Do While rsDetail.EOF = False
      Set rsInventory = New ADODB.Recordset
      rsInventory.Open "SELECT [INV ITEM Id],[INV ITEM Average Cost],[INV ITEM Standard Cost],[INV ITEM Qty On Hand] FROM [INV Items] WHERE [INV ITEM Id] ='" & rsDetail("AR SALED Item ID") & "'", db, adOpenStatic, adLockOptimistic, adCmdText
      'rsInventory.Index = "PrimaryKey"
      'rsInventory.MoveFirst
      'rsInventory.Find "[INV ITEM Id] ='" & rsDetail("AR SALED Item ID") & "'"
      If rsInventory.RecordCount = 0 Then
        'May be a non stock item
      Else
        ' update item cost for this sales detail record
        Dim ItemCost#
        
        ItemCost# = 0
        If CostMethod% = 0 Then
          ItemCost# = IIf(IsNull(rsInventory("INV ITEM Average Cost")), 0, rsInventory("INV ITEM Average Cost"))
          If ItemCost# = 0 Then
            ItemCost# = IIf(IsNull(rsInventory("INV ITEM Standard Cost")), 0, rsInventory("INV ITEM Standard Cost"))
          End If
        Else
          ItemCost# = IIf(IsNull(rsInventory("INV ITEM Standard Cost")), 0, rsInventory("INV ITEM Standard Cost"))
          If ItemCost# = 0 Then
            ItemCost# = IIf(IsNull(rsInventory("INV ITEM Average Cost")), 0, rsInventory("INV ITEM Average Cost"))
          End If
        End If

        rsDetail("AR SALED Item Cost") = ItemCost#
        rsDetail.Update
        ' end of update item cost for this sales detail record
        
        QtySold# = rsDetail("AR SALED Qty") * GetUOMMultiplier(rsDetail("AR SALED Item ID"), rsDetail("AR SALED Units"), db)

        rsInventory("INV ITEM Qty On Hand") = rsInventory("INV ITEM Qty On Hand") - QtySold#
        rsInventory.Update
      End If
      
      rsInventory.Close
      Set rsInventory = Nothing
      
      rsDetail.MoveNext
      'If Err = 3021 Then Exit Do
    Loop
  End If
  ' End of Inventory Updates


  'Dim rsGLTrans As ADODB.Recordset
  'Set rsGLTrans = New ADODB.Recordset
  'rsGLTrans.Open "[GL Transaction]", db, adOpenStatic, adLockOptimistic, adCmdTable

  ' write GL Transaction Header
  Dim refr$
  Dim desc$
  Dim NewNumber&
  'rsGLTrans.AddNew
  
      SQLstatement = "INSERT INTO [GL Transaction]"
      SQLstatement = SQLstatement & " ([GL TRANS Document #],[GL TRANS Type],[GL TRANS Date],"
      SQLstatement = SQLstatement & " [GL TRANS Reference],[GL TRANS Amount],[GL TRANS Posted YN],"
      SQLstatement = SQLstatement & " [GL TRANS Description],[GL TRANS Source],[GL TRANS System Generated])"
    
    'rsGLTrans("GL TRANS Document #") = "INV " & rsSales("AR SALE Ext Document #")
    
    Dim TempStr As String
    
    If rsSales("AR SALE Document Type") = "Service Invoice" Then
      refr$ = "SERINV"
    Else
      refr$ = "Invoice"
    End If
    
    ' gl post date
    
    If PostDate% = 1 Then
      TempStr = Format(Now, "Short Date")
    Else
      TempStr = rsSales("AR SALE Date")
    End If
    
    SQLstatement = SQLstatement & " VALUES ('INV " & rsSales("AR SALE Ext Document #") & "','" & refr$ & "',#" & TempStr & "#,"
    
    If IsNull(rsSales("AR SALE Billing Customer")) Then
        refr$ = rsSales("AR SALE Bill To") & ""
    Else
        refr$ = rsSales("AR SALE Billing Customer") & ""
    End If
    
    desc$ = IIf(IsNull(rsSales("AR SALE Description")), "", rsSales("AR SALE Description"))
    If Len(Trim$(desc$)) = 0 Then
      desc$ = "INV " & rsSales("AR SALE Ext Document #")
    End If
    
      SQLstatement = SQLstatement & "'" & refr$ & "'," & rsSales("AR SALE Total") & ",1,"
      SQLstatement = SQLstatement & "'" & desc$ & "','INV " & rsSales("AR SALE Ext Document #") & "',True)"
      'Debug.Print SQLstatement
      
      db.Execute SQLstatement
    
    'rsGLTrans("GL TRANS Reference") = refr$
    'rsGLTrans("GL TRANS Amount") = rsSales("AR SALE Total")
    'rsGLTrans("GL TRANS Posted YN") = 1
    'rsGLTrans("GL TRANS Description") = desc$
    'rsGLTrans("GL TRANS Source") = "INV " & rsSales("AR SALE Ext Document #")
    'rsGLTrans("GL TRANS System Generated") = True
  'rsGLTrans.Update
  'rsGLTrans.Requery
  'rsGLTrans.MoveLast
  Dim rsGLTrans As ADODB.Recordset
  Set rsGLTrans = New ADODB.Recordset
  rsGLTrans.Open "SELECT [GL TRANS Number] FROM [GL Transaction]WHERE [GL TRANS Document #]='" & "INV " & rsSales("AR SALE Ext Document #") & "'", db, adOpenKeyset, adLockReadOnly, adCmdText
      NewNumber& = rsGLTrans("GL TRANS Number")
  rsGLTrans.Close
  Set rsGLTrans = Nothing
  ' write GL Transaction Detail

  ' update GL
  '  Transaction #1
  '-----------------------------------------------------------------------
  ' Invoice with Payment GL Affected Accounts
  '
  '                  Debit   Credit   Source
  '                  -----   ------   ------
  ' AR                 X              Pref - Sales
  ' CASH               X              Bank - Cash Acct
  ' Discount           X              Pref - Sales
  ' Sales                      X      Item - Inventory or Customer
  ' Sales Tax                  X      Sys Tax
  ' Freight Income             X      Pref - Sales
  ' The following entry is only valid if the payment exceeds the invoice total
  ' AR                         X      Pref - Sales
  '
  ' Notes:
  ' Each Inventory Item is processed and the GL Acct may be retrieved.
  '-----------------------------------------------------------------------

  ' Debits
  ' AR

  If rsSales("AR SALE Total") - rsSales("AR SALE Amount Paid") > 0 Then
      SQLstatement = "INSERT INTO [GL Work Detail]"
      SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
      SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & rsCompany("SYS COM Sales AR Acct") & "" & "'," & rsSales("AR SALE Total") - rsSales("AR SALE Amount Paid") & ",0)"
      db.Execute SQLstatement
'    rsGLWorkDetail.AddNew
'      rsGLWorkDetail("GW TRANSD Number") = NewNumber&
'      rsGLWorkDetail("GW TRANSD Account") = CStr(rsCompany("SYS COM Sales AR Acct"))
'      rsGLWorkDetail("GW TRANSD Debit Amount") = rsSales("AR SALE Total") - rsSales("AR SALE Amount Paid")
'      rsGLWorkDetail("GW TRANSD Credit Amount") = 0
'      rsGLWorkDetail("GW TRANSD Project") = ""
'    rsGLWorkDetail.Update
  End If
  
  ' Cash Receipt
  If rsSales("AR SALE Amount Paid") > 0 Then
      SQLstatement = "INSERT INTO [GL Work Detail]"
      SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
      SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & BankKey$ & "'," & rsSales("AR SALE Amount Paid") & ",0)"
      db.Execute SQLstatement
'    rsGLWorkDetail.AddNew
'      rsGLWorkDetail("GW TRANSD Number") = NewNumber&
'      rsGLWorkDetail("GW TRANSD Account") = BankKey$
'      rsGLWorkDetail("GW TRANSD Debit Amount") = rsSales("AR SALE Amount Paid")
'      rsGLWorkDetail("GW TRANSD Credit Amount") = 0
'      rsGLWorkDetail("GW TRANSD Project") = ""
'    rsGLWorkDetail.Update
  End If
  
  ' Discount Amount
  If rsSales("AR SALE Discount Amount") > 0 Then
      SQLstatement = "INSERT INTO [GL Work Detail]"
      SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
      SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & rsCompany("SYS COM Sales Discount Acct") & "" & "'," & rsSales("AR SALE Discount Amount") & ",0)"
      db.Execute SQLstatement
'    rsGLWorkDetail.AddNew
'      rsGLWorkDetail("GW TRANSD Number") = NewNumber&
'      rsGLWorkDetail("GW TRANSD Account") = rsCompany("SYS COM Sales Discount Acct") & ""
'      rsGLWorkDetail("GW TRANSD Debit Amount") = rsSales("AR SALE Discount Amount")
'      rsGLWorkDetail("GW TRANSD Credit Amount") = 0
'      rsGLWorkDetail("GW TRANSD Project") = ""
'    rsGLWorkDetail.Update
  End If

  ' Credits
  ' Sales Increase
  
  Dim Longer&
  Dim InventoryAcct$
  
  If SalesAcctDefault% = 1 Then
    ' Customer
    If CustomerSalesAcct$ = "" Then ' if no acct then use
      CustomerSalesAcct$ = rsCompany("SYS COM Sales Sales Acct")
    End If
      SQLstatement = "INSERT INTO [GL Work Detail]"
      SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
      SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & CustomerSalesAcct$ & "',0," & rsSales("AR SALE SubTotal") & ")"
      db.Execute SQLstatement
'    rsGLWorkDetail.AddNew
'      rsGLWorkDetail("GW TRANSD Number") = NewNumber&
'      rsGLWorkDetail("GW TRANSD Account") = CustomerSalesAcct$
'      rsGLWorkDetail("GW TRANSD Debit Amount") = 0
'      rsGLWorkDetail("GW TRANSD Credit Amount") = rsSales("AR SALE SubTotal")
'      rsGLWorkDetail("GW TRANSD Project") = ""
'    rsGLWorkDetail.Update
  Else
    ' item
    Longer& = 0
    If rsDetail.RecordCount = 0 Then
      ' no detail data
    Else
    rsDetail.MoveFirst
      ' may have detail data
      Do While rsDetail.EOF = False
          InventoryAcct$ = NZ(rsDetail("AR SALED Posting Account"))
          If InventoryAcct$ = "" Then
            If rsDetail("AR SALED Row Type") = "N" Then ' process non-stock
              InventoryAcct$ = rsCompany("SYS COM Sales Misc Acct")
            Else
              'rsInventory.Index = "PrimaryKey"
              'rsInventory.Seek rsDetail("AR SALED Item ID")
                Set rsInventory = New ADODB.Recordset
                rsInventory.Open "SELECT [INV ITEM Id],[INV ITEM Sales Account] FROM [INV Items] WHERE [INV ITEM Id]='" & rsDetail("AR SALED Item ID") & "'", db, adOpenStatic, adLockOptimistic, adCmdText
             
              'rsInventory.MoveFirst
              'rsInventory.Find "[INV ITEM Id]='" & rsDetail("AR SALED Item ID") & "'"
              If rsInventory.RecordCount = 0 Then
                InventoryAcct$ = rsCompany("SYS COM Sales Sales Acct")
              Else
                InventoryAcct$ = IIf(IsNull(rsInventory("INV ITEM Sales Account")), rsCompany("SYS COM Sales Sales Acct"), rsInventory("INV ITEM Sales Account"))
              End If
              
                rsInventory.Close
                Set rsInventory = Nothing
            End If
          End If
            SQLstatement = "INSERT INTO [GL Work Detail]"
            SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
            SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & InventoryAcct$ & "',0," & rsDetail("AR SALED Item Total") & ")"
            db.Execute SQLstatement
          'rsGLWorkDetail.AddNew
          '  rsGLWorkDetail("GW TRANSD Number") = NewNumber&
          '  rsGLWorkDetail("GW TRANSD Account") = InventoryAcct$
          '  rsGLWorkDetail("GW TRANSD Debit Amount") = 0
          '  rsGLWorkDetail("GW TRANSD Credit Amount") = rsDetail("AR SALED Item Total")
          '  rsGLWorkDetail("GW TRANSD Project") = ""
          'rsGLWorkDetail.Update

        rsDetail.MoveNext
        'If Err = 3021 Then Exit Do
      Loop
    End If
  End If
  
  ' Sales Tax
  Dim GLTaxAcct$
  Dim TaxGroup$
  Dim TaxPercent#
  Dim TaxID$
  Dim Taxtotal#

  If rsSales("AR SALE Sales Tax") > 0 Then
    GLTaxAcct$ = ""
    TaxGroup$ = IIf(IsNull(rsSales("AR SALE Tax Group")), rsCompany("SYS COM Sales Sales Tax"), rsSales("AR SALE Tax Group"))
    Dim rsTaxGroupDetail As ADODB.Recordset
    
    Set rsTaxGroupDetail = New ADODB.Recordset
    rsTaxGroupDetail.Open "SELECT * FROM [SYS Tax Group Detail] where [SYS TAXGRPD Group ID] = '" & TaxGroup$ & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    
    If rsTaxGroupDetail.RecordCount = 0 Then
        rsTaxGroupDetail.MoveFirst
    Else
      Do While rsTaxGroupDetail.EOF = False
        TaxPercent# = 0
        TaxID$ = rsTaxGroupDetail("SYS TAXGRPD Tax ID")
        TaxPercent# = LookRecord("[SYS Tax Percent]", "[SYS Tax]", db, "[SYS Tax ID] = '" & TaxID$ & "'")
        GLTaxAcct$ = NZ(LookRecord("[SYS Tax Account]", "[SYS Tax]", db, "[SYS Tax ID] = '" & TaxID$ & "'"))
        
        If GLTaxAcct$ = "" Then
          GLTaxAcct$ = rsCompany("SYS COM Sales Sales Acct")
        End If
        Taxtotal# = Round(rsSales("AR SALE Taxable Subtotal") * (TaxPercent# / 100))
        SQLstatement = "INSERT INTO [GL Work Detail]"
        SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
        SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & GLTaxAcct$ & "',0," & Taxtotal# & ")"
        db.Execute SQLstatement
        'rsGLWorkDetail.AddNew
        '  rsGLWorkDetail("GW TRANSD Number") = NewNumber&
        '  rsGLWorkDetail("GW TRANSD Account") = GLTaxAcct$
        '  rsGLWorkDetail("GW TRANSD Debit Amount") = 0
        '  rsGLWorkDetail("GW TRANSD Credit Amount") = Round(Taxtotal#)
        '  rsGLWorkDetail("GW TRANSD Project") = ""
        'rsGLWorkDetail.Update

        rsTaxGroupDetail.MoveNext
      Loop
      rsTaxGroupDetail.Close
      Set rsTaxGroupDetail = Nothing
    End If
  End If
  ' Sales Tax
  
  ' Freight Income
  If rsSales("AR SALE Freight") > 0 Then
    SQLstatement = "INSERT INTO [GL Work Detail]"
    SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
    SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & rsCompany("SYS COM Sales Freight Acct") & "" & "',0," & rsSales("AR SALE Freight") & ")"
    db.Execute SQLstatement
    'rsGLWorkDetail.AddNew
    '  rsGLWorkDetail("GW TRANSD Number") = NewNumber&
    '  rsGLWorkDetail("GW TRANSD Account") = rsCompany("SYS COM Sales Freight Acct")
    '  rsGLWorkDetail("GW TRANSD Debit Amount") = 0
    '  rsGLWorkDetail("GW TRANSD Credit Amount") = rsSales("AR SALE Freight")
    '  rsGLWorkDetail("GW TRANSD Project") = ""
    'rsGLWorkDetail.Update
    
  End If

  ' update GL
  '  Transaction #2
  ' this transaction is done in reverse order so we can accumulate the
  '   credits and generate a corresponding debit and header entry(total)
  '-----------------------------------------------------------------------
  ' Invoice with Payment GL Affected Accounts - COGS
  '
  '                  Debit   Credit   Source
  '                  -----   ------   ------
  ' Inventory                  X      Item - Inventory
  ' COGS               X              Pref - Sales
  '
  ' Notes:
  ' Each Inventory Item is processed and the GL Acct is Retrieved.
  '-----------------------------------------------------------------------
  ' we need a new GL Number
  

  Dim COGS@
  Dim InventoryCost@
  Dim COGSAccount$

  COGS@ = 0
    
  ' item Inventory entry
  Longer& = 0
  
  
  If rsDetail.RecordCount = 0 Then
    ' no detail data
  Else
  rsDetail.MoveFirst
    ' may have detail data
    Do While rsDetail.EOF = False
        InventoryAcct$ = ""
        COGSAccount$ = ""
        InventoryCost@ = 0
        If rsDetail("AR SALED Row Type") = "N" Then ' process non-stock
          InventoryAcct$ = rsCompany("SYS COM Sales Misc Acct")
          COGSAccount$ = rsCompany("SYS COM Sales COGS Acct")
        Else
          Set rsInventory = New ADODB.Recordset
          rsInventory.Open "SELECT [INV ITEM Id],[INV ITEM Inventory Account],[INV ITEM Cost Of Sales Account],[INV ITEM Average Cost] FROM [INV Items] WHERE [INV ITEM Id]='" & rsDetail("AR SALED Item ID") & "'", db, adOpenStatic, adLockOptimistic, adCmdText
          'rsInventory.Index = "PrimaryKey"
          'rsInventory.MoveFirst
          'rsInventory.Find "[INV ITEM Id]='" & rsDetail("AR SALED Item ID") & "'"
          If rsInventory.RecordCount = 0 Then
            InventoryAcct$ = rsCompany("SYS COM Sales Inventory Acct")
            COGSAccount$ = rsCompany("SYS COM Sales COGS Acct")
          Else
            InventoryAcct$ = IIf(IsNull(rsInventory("INV ITEM Inventory Account")), rsCompany("SYS COM Sales Inventory Acct"), rsInventory("INV ITEM Inventory Account"))
            COGSAccount$ = IIf(IsNull(rsInventory("INV ITEM Cost Of Sales Account")), rsCompany("SYS COM Sales COGS Acct"), rsInventory("INV ITEM Cost Of Sales Account"))
            InventoryCost@ = IIf(IsNull(rsInventory("INV ITEM Average Cost")), 0, rsInventory("INV ITEM Average Cost"))
          End If
            rsInventory.Close
            Set rsInventory = Nothing
        End If
          
        If InventoryCost@ = 0 Then InventoryCost@ = IIf(IsNull(rsDetail("AR SALED Item Cost")), 0, rsDetail("AR SALED Item Cost"))
          
        ' Use Inventory Rounding
        Dim TempCost#
        TempCost# = InventoryCost@
        InventoryCost@ = Round(TempCost#)

        ' Don't do a GetUOMMultiplier on non-stock items
        If rsDetail("AR SALED Row Type") = "N" Then
          QtySold# = rsDetail("AR SALED Qty")
        Else
          QtySold# = rsDetail("AR SALED Qty") * GetUOMMultiplier(rsDetail("AR SALED Item ID"), rsDetail("AR SALED Units"), db)
        End If

        InventoryCost@ = InventoryCost@ * QtySold#    'tbARSalesDetail("AR SALED Qty")

        InventoryCost@ = DropAllBut2(InventoryCost@)

        If InventoryCost@ = 0 Then
        Else
          COGS@ = COGS@ + InventoryCost@
            SQLstatement = "INSERT INTO [GL Work Detail]"
            SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
            SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & InventoryAcct$ & "',0," & InventoryCost@ & ")"
            db.Execute SQLstatement
          'rsGLWorkDetail.AddNew
          '  rsGLWorkDetail("GW TRANSD Number") = NewNumber&
          '  rsGLWorkDetail("GW TRANSD Account") = InventoryAcct$
          '  rsGLWorkDetail("GW TRANSD Debit Amount") = 0
          '  rsGLWorkDetail("GW TRANSD Credit Amount") = InventoryCost@
          '  rsGLWorkDetail("GW TRANSD Project") = ""
          'rsGLWorkDetail.Update
          
          ' debit for COGS
            SQLstatement = "INSERT INTO [GL Work Detail]"
            SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
            SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & COGSAccount$ & "'," & InventoryCost@ & ",0)"
            db.Execute SQLstatement
          'rsGLWorkDetail.AddNew
          '  rsGLWorkDetail("GW TRANSD Number") = NewNumber&
          '  rsGLWorkDetail("GW TRANSD Account") = COGSAccount$    'rsCompany("SYS COM Sales COGS Acct")
          '  rsGLWorkDetail("GW TRANSD Debit Amount") = InventoryCost@
          '  rsGLWorkDetail("GW TRANSD Credit Amount") = 0
          '  rsGLWorkDetail("GW TRANSD Project") = ""
          'rsGLWorkDetail.Update
    
        End If
      rsDetail.MoveNext
      'If Err = 3021 Then Exit Do
    Loop
  End If

  ' post GL entry for this Invoice
  Dim Success%
  'Success% = PostGLWorkDetail(TranDate, NewNumber&)
  Success% = PostGLWorkDetail(TranDate, NewNumber&, db)
  If Success% = False Then
    MsgBox "An error occurred writing GL Transaction!", , "Error"
    PostInvoice = False
    Exit Function
  End If

PostInvoice_Exit:

  PostInvoice = True
'  rsGLWorkDetail.Close
'  Set rsGLWorkDetail = Nothing
  rsDetail.Close
  Set rsDetail = Nothing
'  rsTaxGroupDetail.Close
'  Set rsTaxGroupDetail = Nothing
'  rsGLTrans.Close
'  Set rsGLTrans = Nothing
  rsCompany.Close
  Set rsCompany = Nothing
  rsSales.Close
  Set rsSales = Nothing
'  rsARPaymentHeader.Close    '<<<==============check code
'  Set rsARPaymentHeader = Nothing
'  rsARCross.Close            '<<<==============check code
'  Set rsARCross = Nothing
  rsCustomer.Close
  Set rsCustomer = Nothing
'  rsCompany.Close            '<<<==============check code
'  Set rsCompany = Nothing
'  rsInventory.Close
'  Set rsInventory = Nothing
If DBload = True Then
  db.Close
  Set db = Nothing
End If

  Exit Function

PostInvoice_Error:
  Call ErrorLog("Sales Module", "PostInvoice", Now, Err.Number, Err.Description, intShowError, db)
  PostInvoice = False
  Resume Next
'  rsGLWorkDetail.Close
'  Set rsGLWorkDetail = Nothing
  rsDetail.Close
  Set rsDetail = Nothing
'  rsTaxGroupDetail.Close
'  Set rsTaxGroupDetail = Nothing
'  rsGLTrans.Close
'  Set rsGLTrans = Nothing
  rsCompany.Close
  Set rsCompany = Nothing
  rsSales.Close
  Set rsSales = Nothing
'  rsARPaymentHeader.Close
'  Set rsARPaymentHeader = Nothing
'  rsARCross.Close
'  Set rsARCross = Nothing
  rsCustomer.Close
  Set rsCustomer = Nothing
'  rsCompany.Close
'  Set rsCompany = Nothing
'  rsInventory.Close
'  Set rsInventory = Nothing
If DBload = True Then
  db.Close
  Set db = Nothing
End If
  Exit Function

UnableToPostHere:
  PostInvoice = False
'  rsGLWorkDetail.Close
'  Set rsGLWorkDetail = Nothing
  rsDetail.Close
  Set rsDetail = Nothing
'  rsTaxGroupDetail.Close
'  Set rsTaxGroupDetail = Nothing
'  rsGLTrans.Close
'  Set rsGLTrans = Nothing
  rsCompany.Close
  Set rsCompany = Nothing
  rsSales.Close
  Set rsSales = Nothing
'  rsARPaymentHeader.Close
'  Set rsARPaymentHeader = Nothing
'  rsARCross.Close
'  Set rsARCross = Nothing
  rsCustomer.Close
  Set rsCustomer = Nothing
'  rsCompany.Close
'  Set rsCompany = Nothing
'  rsInventory.Close
'  Set rsInventory = Nothing
If DBload = True Then
  db.Close
  Set db = Nothing
End If

  Exit Function
  
'  rsGLWorkDetail.Close
'  Set rsGLWorkDetail = Nothing
  rsDetail.Close
  Set rsDetail = Nothing
'  rsTaxGroupDetail.Close
'  Set rsTaxGroupDetail = Nothing
'  rsGLTrans.Close
'  Set rsGLTrans = Nothing
  rsCompany.Close
  Set rsCompany = Nothing
  rsSales.Close
  Set rsSales = Nothing
'  rsARPaymentHeader.Close
'  Set rsARPaymentHeader = Nothing
'  rsARCross.Close
'  Set rsARCross = Nothing
  rsCustomer.Close
  Set rsCustomer = Nothing
'  rsCompany.Close
'  Set rsCompany = Nothing
'  rsInventory.Close
'  Set rsInventory = Nothing
If DBload = True Then
  db.Close
  Set db = Nothing
End If
End Function

Function PostReturn(DocumentKey&, intShowError As Integer, db As ADODB.Connection) As Integer

  ''On Error GoTo PostReturn_Error

  Dim msg$
  Dim title$
  
  'Dim db As ADODB.Connection
  'Set db = New ADODB.Connection
  'db.CursorLocation = adUseServer
  'db.Open gblADOProvider
  
  Dim rsCompany As ADODB.Recordset
  Set rsCompany = New ADODB.Recordset
  rsCompany.Open "SELECT [SYS COM GL Post By Date],[SYS COM Sales Acct Default]," & _
  "[SYS COM Inventory Cost Method Last YN],[SYS COM Sales AR Acct],[SYS COM Sales Discount Acct]," & _
  "[SYS COM Sales Sales Acct],[SYS COM Sales Misc Acct],[SYS COM Sales Sales Tax]," & _
  "[SYS COM Sales Freight Acct],[SYS COM Sales Return Acct],[SYS COM Sales COGS Acct],[SYS COM Sales Inventory Acct] " & _
  "FROM [SYS Company]", db, adOpenKeyset, adLockOptimistic, adCmdText
    
  rsCompany.MoveFirst

  'Dim rsGLWorkDetail As ADODB.Recordset
  'Set rsGLWorkDetail = New ADODB.Recordset
  'rsGLWorkDetail.Open "[GL Work Detail]", db, adOpenStatic, adLockOptimistic, adCmdTable

  Dim rsSales As ADODB.Recordset
  Set rsSales = New ADODB.Recordset '
  rsSales.Open "SELECT [AR SALE Date],[AR SALE Customer ID],[AR SALE Total],[AR SALE Ext Document #]," & _
  "[AR SALE Check Acct ID],[AR SALE Billing Customer],[AR SALE Bill To],[AR SALE Description]," & _
  "[AR SALE Discount Amount],[AR SALE Sales Tax],[AR SALE Tax Group],[AR SALE Taxable Subtotal]," & _
  "[AR SALE Freight],[AR SALE Document #],[AR SALE Amount Paid],[AR SALE Document Type]," & _
  "[AR SALE SubTotal] FROM [AR Sales] WHERE [AR SALE Document #]=" & DocumentKey&, db, adOpenKeyset, adLockOptimistic, adCmdText
  'rsSales.Open "[AR Sales]", db, adOpenStatic, adLockOptimistic, adCmdTable
  'rsSales.Index = "PrimaryKey"
  'rsSales.MoveFirst
  'rsSales.Find "[AR SALE Document #]='" & DocumentKey& & "'"

  'Post by 1-System Date or 2-Transaction Date?
  Dim PostDate%
  PostDate% = rsCompany![SYS COM GL Post By Date]

  Dim InvoiceType$
  InvoiceType$ = "Return"
  
  'Set Post Date
  Dim TranDate As Variant
  If PostDate% = 1 Then
    TranDate = DateValue(Format(Now, "Short Date"))
  Else
    TranDate = DateValue(rsSales![AR SALE Date])
  End If
  

  'Verify period can be posted to
  'Send TranDate
  'Return PeriodToPost and PeriodClose
  Dim PeriodToPost%
  Dim PeriodClosed%
  
  Call VerifyPeriod(TranDate, PeriodToPost%, PeriodClosed%, db)
  If PeriodClosed% = True Then
    MsgBox "Unable to post transaction to a closed period.", , "Post Return Error"
    GoTo UnableToPostReturnHere
    Exit Function
  End If

  ''On Error GoTo PostReturns_Error

  ' clear any GL Work records
  'Dim cmdtemp As ADODB.Recordset
  'cmdtemp.Open "DELETE DISTINCTROW * FROM [GL Work Detail]", db, , , adCmdText
  db.Execute "DELETE DISTINCTROW * FROM [GL Work Detail]", , adCmdText
  'cmdtemp.Close
  'Set cmdtemp = Nothing

  Dim rsCustomer As ADODB.Recordset
  Set rsCustomer = New ADODB.Recordset
  'rsCustomer.Open "SELECT * FROM [AR Customer] where [AR CUST Customer ID] = '" & rsSales![AR SALE Customer ID] & "'", db, adOpenStatic, adLockOptimistic, adCmdText
  rsCustomer.Open "SELECT [AR CUST Customer ID],[AR CUST Payments YTD],[AR CUST Payments Lifetime],[AR CUST Sales YTD]," & _
  "[AR CUST Sales Lifetime],[AR CUST Invoices Lifetime],[AR CUST Invoices YTD],[AR CUST Financial Period 1]," & _
  "[AR CUST Highest Balance] FROM [AR Customer] WHERE [AR CUST Customer ID] = '" & rsSales("AR SALE Customer ID") & "'", db, adOpenStatic, adLockOptimistic, adCmdText

  ' if we are to post by customer get the Customer Sales Acct
  Dim SalesAcctDefault%
  Dim CustomerSalesAcct$

  '1-Post usting customer accounts; 2-Post using item accounts
  SalesAcctDefault% = rsCompany![SYS COM Sales Acct Default]

  CustomerSalesAcct$ = ""
  If SalesAcctDefault% = 1 Then
    CustomerSalesAcct$ = IIf(IsNull(rsCustomer![AR CUST Sales Account]), "", rsCustomer![AR CUST Sales Account])
  End If

  Dim CurrentBalance@

    rsCustomer![AR CUST Sales YTD] = rsCustomer![AR CUST Sales YTD] - rsSales![AR SALE Total]
    rsCustomer![AR CUST Sales Lifetime] = rsCustomer![AR CUST Sales Lifetime] - rsSales![AR SALE Total]
    ' Update current Balance - if not Paid in full
    CurrentBalance@ = IIf(IsNull(rsCustomer![AR CUST Financial Period 1]), 0, rsCustomer![AR CUST Financial Period 1])
    CurrentBalance@ = CurrentBalance@ - rsSales![AR SALE Total]
    rsCustomer![AR CUST Financial Period 1] = CurrentBalance@
  rsCustomer.Update
  rsCustomer.Close
  Set rsCustomer = Nothing
 

  '--------------------------------------------------

  ' Inventory Updates
  Dim rsInventory As ADODB.Recordset
  Dim rsDetail As ADODB.Recordset
  
  Set rsInventory = New ADODB.Recordset
  Set rsDetail = New ADODB.Recordset
  rsInventory.Open "SELECT [INV ITEM Id],[INV ITEM Average Cost],[INV ITEM Standard Cost],[INV ITEM Qty On Hand],[INV ITEM Inventory Account],[INV ITEM Cost Of Sales Account] FROM [INV Items]", db, adOpenStatic, adLockOptimistic, adCmdText
  rsDetail.Open "SELECT * FROM [AR Sales Detail] where [AR SALED Document #] = " & rsSales![AR SALE Document #], db, adOpenStatic, adLockOptimistic, adCmdText
  '
  '

  'rsDetail.MoveLast
  rsDetail.MoveFirst
  Dim QtySold#
  Dim CostMethod%
  '-1-Use Last Cost Method; 0-Use Average Costing

  If rsDetail.RecordCount = 0 Then
    'No detail data
  Else
    Do While Not rsDetail.EOF
      'rsInventory.Index = "PrimaryKey"
      rsInventory.MoveFirst
      rsInventory.Find "[INV ITEM Id]='" & rsDetail![AR SALED Item ID] & "'"
      If rsInventory.EOF = True Then
        ' May be a non-stock item
      Else
        
        ' update item cost for this sales detail record
        Dim ItemCost#
        
        ItemCost# = 0
        If CostMethod% = 0 Then
          ItemCost# = IIf(IsNull(rsInventory("INV ITEM Average Cost")), 0, rsInventory("INV ITEM Average Cost"))
          If ItemCost# = 0 Then
            ItemCost# = IIf(IsNull(rsInventory("INV ITEM Standard Cost")), 0, rsInventory("INV ITEM Standard Cost"))
          End If
        Else
          ItemCost# = IIf(IsNull(rsInventory("INV ITEM Standard Cost")), 0, rsInventory("INV ITEM Standard Cost"))
          If ItemCost# = 0 Then
            ItemCost# = IIf(IsNull(rsInventory("INV ITEM Average Cost")), 0, rsInventory("INV ITEM Average Cost"))
          End If
        End If

        rsDetail("AR SALED Item Cost") = ItemCost#
        rsDetail.Update
        ' end of update item cost for this sales detail record

        QtySold# = rsDetail("AR SALED Qty") * GetUOMMultiplier(rsDetail("AR SALED Item ID"), rsDetail("AR SALED Units"), db)

        rsInventory("INV ITEM Qty On Hand") = rsInventory("INV ITEM Qty On Hand") + QtySold#    'tbARSalesDetail("AR SALED Qty")
        rsInventory.Update
      End If
      
      rsDetail.MoveNext
      'If Err = 3021 Then Exit Do
    Loop
  End If
  ' End of Inventory Updates
  '--------------------------------------------------
  ' AR Cross Payment and AR Payment Header

  Dim PaymentID&
  'Dim rsARPaymentHeader As ADODB.Recordset
  'Set rsARPaymentHeader = New ADODB.Recordset
  'rsARPaymentHeader.Open "[AR Payment Header]", db, adOpenStatic, adLockOptimistic, adCmdTable
    
    'Dim rsARCross As ADODB.Recordset
    'Set rsARCross = New ADODB.Recordset
    'rsARCross.Open "[AR Payment Invoice Cross Reference]", db, adOpenStatic, adLockOptimistic, adCmdTable
    
    '  write Payment Header
    'rsARPaymentHeader.AddNew
    '  rsARPaymentHeader("AR PAY Type") = "Return"
    '  rsARPaymentHeader("AR PAY Check No") = "RET " & rsSales("AR SALE Ext Document #")
    '  rsARPaymentHeader("AR PAY Customer No") = rsSales("AR SALE Customer ID") & ""
    '  rsARPaymentHeader("AR PAY Transaction Date") = rsSales("AR SALE Date")
    '  rsARPaymentHeader("AR PAY Amount") = rsSales("AR SALE Total")
    '  rsARPaymentHeader("AR PAY UnApplied Amount") = rsSales("AR SALE Total")  'Cannot have unapplied amounts here
    '  rsARPaymentHeader("AR PAY Bank Account") = "None"
    '  rsARPaymentHeader("AR PAY Status") = "Posted"
    '  rsARPaymentHeader("AR PAY NSF") = False
    '  rsARPaymentHeader("AR PAY Posted YN") = True
    '  rsARPaymentHeader("AR PAY Cleared") = False
    'rsARPaymentHeader.Update
      
      SQLstatement = "INSERT INTO [AR Payment Header]"
      SQLstatement = SQLstatement & " ([AR PAY Type],[AR PAY Check No],[AR PAY Customer No],[AR PAY Transaction Date],[AR PAY Amount],[AR PAY UnApplied Amount],[AR PAY Bank Account],[AR PAY Status],[AR PAY Posted YN],[AR PAY NSF],[AR PAY Cleared])"
      SQLstatement = SQLstatement & " VALUES ('Return','RET " & rsSales("AR SALE Ext Document #") & "','" & CStr(rsSales("AR SALE Customer ID")) & "',#" & rsSales("AR SALE Date") & "#," & rsSales("AR SALE Total") & "," & rsSales("AR SALE Total") & ",'None','Posted',True,False,False)"
      db.Execute SQLstatement
      
      Dim rsARPaymentHeader As ADODB.Recordset
      Set rsARPaymentHeader = New ADODB.Recordset
      rsARPaymentHeader.Open "SELECT [AR PAY ID] FROM [AR Payment Header] WHERE [AR PAY Check No]='RET " & rsSales("AR SALE Ext Document #") & "'", db, adOpenKeyset, adLockReadOnly, adCmdText
      If rsARPaymentHeader.RecordCount > 1 Then
        rsARPaymentHeader.MoveLast
        PaymentID& = rsARPaymentHeader("AR PAY ID")
      Else
        PaymentID& = rsARPaymentHeader("AR PAY ID")
      End If
      rsARPaymentHeader.Close
      Set rsARPaymentHeader = Nothing
    
    ' end of write payment header

'  End If
  ' end of write AR Cross reference Record

  'Write GL Transaction Header
  'Dim rsGLTrans As ADODB.Recordset
  'Set rsGLTrans = New ADODB.Recordset
  'rsGLTrans.Open "[GL Transaction]", db, adOpenStatic, adLockOptimistic, adCmdTable

  Dim refr$
  Dim desc$
  Dim NewNumber&
  'rsGLTrans.AddNew
      SQLstatement = "INSERT INTO [GL Transaction]"
      SQLstatement = SQLstatement & " ([GL TRANS Document #],[GL TRANS Type],[GL TRANS Date],"
      SQLstatement = SQLstatement & " [GL TRANS Reference],[GL TRANS Amount],[GL TRANS Posted YN],"
      SQLstatement = SQLstatement & " [GL TRANS Description],[GL TRANS Source],[GL TRANS System Generated])"

    'rsGLTrans("GL TRANS Document #") = "RET " & rsSales("AR SALE Ext Document #")
    'rsGLTrans("GL TRANS Type") = "Return"

    If PostDate% = 1 Then
      TempStr = Format(Now, "Short Date")
    Else
      TempStr = rsSales("AR SALE Date")
    End If
    
    SQLstatement = SQLstatement & " VALUES ('RET " & rsSales("AR SALE Ext Document #") & "','Return',#" & TempStr & "#,"

    If IsNull(rsSales("AR SALE Billing Customer")) Then
        refr$ = rsSales("AR SALE Bill To") & ""
    Else
        refr$ = rsSales("AR SALE Billing Customer") & ""
    End If
    
    desc$ = IIf(IsNull(rsSales("AR SALE Description")), "", rsSales("AR SALE Description"))
    If Len(Trim$(desc$)) = 0 Then
      desc$ = "RET " & rsSales("AR SALE Ext Document #")
    End If

      SQLstatement = SQLstatement & "'" & refr$ & "'," & rsSales("AR SALE Total") & ",True,"
      SQLstatement = SQLstatement & "'" & desc$ & "','RET " & rsSales("AR SALE Ext Document #") & "',True)"
      'Debug.Print SQLstatement
      
      db.Execute SQLstatement
      
    'rsGLTrans("GL TRANS Reference") = refr$
    'rsGLTrans("GL TRANS Amount") = rsSales("AR SALE Total")
    'rsGLTrans("GL TRANS Posted YN") = True
    'desc$ = IIf(IsNull(rsSales("AR SALE Description")), "", rsSales("AR SALE Description"))
    'If Len(Trim$(desc$)) = 0 Then
    '  desc$ = "RET " & rsSales("AR SALE Ext Document #")
    'End If
    'rsGLTrans("GL TRANS Description") = desc$
    'rsGLTrans("GL TRANS Source") = "RET " & rsSales("AR SALE Ext Document #")
  'rsGLTrans.Update
  '  NewNumber& = rsGLTrans("GL TRANS Number")
  Dim rsGLTrans As ADODB.Recordset
  Set rsGLTrans = New ADODB.Recordset
  rsGLTrans.Open "SELECT [GL TRANS Number] FROM [GL Transaction] WHERE [GL TRANS Document #]='" & "RET " & rsSales("AR SALE Ext Document #") & "'", db, adOpenKeyset, adLockReadOnly, adCmdText
      NewNumber& = rsGLTrans("GL TRANS Number")
  rsGLTrans.Close
  Set rsGLTrans = Nothing

  ' update GL
  '  Transaction #1
  '-----------------------------------------------------------------------
  ' Returns GL Affected Accounts
  '
  '                  Debit   Credit   Source
  '                  -----   ------   ------
  ' Sales Returns      X              Pref - Sales - gSalesReturnAcct$
  ' Sales Tax          X              Sys Tax
  ' Freight Income     X              Pref - Sales
  ' AR                         X      Pref - Sales
  ' Discount                   X      Pref - Sales
  ' CASH                       X      Bank - Cash Acct xxx 11/21/95
  '
  ' Notes:
  ' Each Inventory Item is processed and the GL Acct may be retrieved.
  '-----------------------------------------------------------------------

  ' Debits
  ' Sales Returns
      SQLstatement = "INSERT INTO [GL Work Detail]"
      SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
      SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & rsCompany("SYS COM Sales Return Acct") & "" & "'," & rsSales("AR SALE SubTotal") & ",0)"
      db.Execute SQLstatement
  
  'rsGLWorkDetail.AddNew
  '  rsGLWorkDetail("GW TRANSD Number") = NewNumber&
  '  rsGLWorkDetail("GW TRANSD Account") = rsCompany("SYS COM Sales Return Acct")
  '  rsGLWorkDetail("GW TRANSD Debit Amount") = rsSales("AR SALE SubTotal")
  '  rsGLWorkDetail("GW TRANSD Credit Amount") = 0
  '  rsGLWorkDetail("GW TRANSD Project") = ""
  'rsGLWorkDetail.Update
  
  Dim TotalTaxPercent#
  Dim GLTaxAcct$
  Dim TaxGroup$
  Dim TaxID$
  Dim TaxPercent#
  Dim Taxtotal#
  
  ' Sales Tax - Fixed
  If rsSales("AR SALE Sales Tax") > 0 Then
    TotalTaxPercent# = 0
    GLTaxAcct$ = ""
    TaxGroup$ = IIf(IsNull(rsSales("AR SALE Tax Group")), rsCompany("SYS COM Sales Sales Tax"), rsSales("AR SALE Tax Group"))
    
    Dim rsTaxGroupDetail As ADODB.Recordset
    Set rsTaxGroupDetail = New ADODB.Recordset
    rsTaxGroupDetail.Open "SELECT * FROM [SYS TAX Group Detail] where [SYS TAXGRPD Group ID] = '" & TaxGroup$ & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    'rsTaxGroupDetail.MoveLast
    If rsTaxGroupDetail.RecordCount = 0 Then
    Else
      rsTaxGroupDetail.MoveFirst
      Do While rsTaxGroupDetail.EOF = False
        
        TaxPercent# = 0
        TaxID$ = rsTaxGroupDetail("SYS TAXGRPD Tax ID")
        TaxPercent# = LookRecord("[SYS Tax Percent]", "[SYS Tax]", db, "[SYS Tax ID] = '" & TaxID$ & "'")
        GLTaxAcct$ = LookRecord("[SYS Tax Account]", "[SYS Tax]", db, "[SYS Tax ID] = '" & TaxID$ & "'")
        
        If GLTaxAcct$ = "" Then
          GLTaxAcct$ = rsCompany("SYS COM Sales Sales Acct")
        End If
        Taxtotal# = Round(rsSales("AR SALE Taxable Subtotal") * (TaxPercent# / 100))
        SQLstatement = "INSERT INTO [GL Work Detail]"
        SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
        SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & GLTaxAcct$ & "'," & Taxtotal# & ",0)"
        db.Execute SQLstatement

        'rsGLWorkDetail.AddNew
        '  rsGLWorkDetail("GW TRANSD Number") = NewNumber&
        '  rsGLWorkDetail("GW TRANSD Account") = GLTaxAcct$
        '  rsGLWorkDetail("GW TRANSD Debit Amount") = Round(rsSales("AR SALE Taxable Subtotal") * (TaxPercent# / 100))
        '  rsGLWorkDetail("GW TRANSD Credit Amount") = 0
        '  rsGLWorkDetail("GW TRANSD Project") = ""
        'rsGLWorkDetail.Update

        rsTaxGroupDetail.MoveNext
      Loop
      rsTaxGroupDetail.Close
      Set rsTaxGroupDetail = Nothing
    End If
  End If
  ' Sales Tax - Fixed

  ' Freight Income
  If rsSales("AR SALE Freight") > 0 Then
    SQLstatement = "INSERT INTO [GL Work Detail]"
    SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
    SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & rsCompany("SYS COM Sales Freight Acct") & "" & "'," & rsSales("AR SALE Freight") & ",0)"
    db.Execute SQLstatement
    'rsGLWorkDetail.AddNew
    '  rsGLWorkDetail("GW TRANSD Number") = NewNumber&
    '  rsGLWorkDetail("GW TRANSD Account") = rsCompany("SYS COM Sales Freight Acct")
    '  rsGLWorkDetail("GW TRANSD Debit Amount") = rsSales("AR SALE Freight")
    '  rsGLWorkDetail("GW TRANSD Credit Amount") = 0
    '  rsGLWorkDetail("GW TRANSD Project") = ""
    'rsGLWorkDetail.Update

  End If

  ' Credits
  ' AR
  If rsSales("AR SALE Total") > 0 Then
    SQLstatement = "INSERT INTO [GL Work Detail]"
    SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
    SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & rsCompany("SYS COM Sales AR Acct") & "" & "',0," & rsSales("AR SALE Total") & ")"
    db.Execute SQLstatement
    'rsGLWorkDetail.AddNew
    '  rsGLWorkDetail("GW TRANSD Number") = NewNumber&
    '  rsGLWorkDetail("GW TRANSD Account") = rsCompany("SYS COM Sales AR Acct")
    '  rsGLWorkDetail("GW TRANSD Debit Amount") = 0
    '  rsGLWorkDetail("GW TRANSD Credit Amount") = rsSales("AR SALE Total")
    '  rsGLWorkDetail("GW TRANSD Project") = ""
    'rsGLWorkDetail.Update
  End If
  
  ' Discount Amount
  If rsSales("AR SALE Discount Amount") > 0 Then
    SQLstatement = "INSERT INTO [GL Work Detail]"
    SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
    SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & rsCompany("SYS COM Sales Discount Acct") & "" & "',0," & rsSales("AR SALE Discount Amount") & ")"
    db.Execute SQLstatement
    'rsGLWorkDetail.AddNew
    '  rsGLWorkDetail("GW TRANSD Number") = NewNumber&
    '  rsGLWorkDetail("GW TRANSD Account") = rsCompany("SYS COM Sales Discount Acct")
    '  rsGLWorkDetail("GW TRANSD Debit Amount") = 0
    '  rsGLWorkDetail("GW TRANSD Credit Amount") = rsSales("AR SALE Discount Amount")
    '  rsGLWorkDetail("GW TRANSD Project") = ""
    'rsGLWorkDetail.Update
  End If
  
  ' update GL
  '  Transaction #2
  '-----------------------------------------------------------------------
  ' Invoice with Payment GL Affected Accounts - COGS
  '
  '                  Debit   Credit   Source
  '                  -----   ------   ------
  ' Inventory          X              Item - Inventory
  ' COGS                        X     Pref - Sales
  '
  ' Notes:
  ' Each Inventory Item is processed and the GL Acct is Retrieved.
  '-----------------------------------------------------------------------

  Dim COGS@
  Dim InventoryCost@
  Dim COGSAccount$
  
  COGS@ = 0
  

  ' item Inventory entry
  Dim InventoryAcct$

  'rsDetail.MoveLast

  If rsDetail.RecordCount = 0 Then
    'No line items
  Else
  rsDetail.MoveFirst
    Do While rsDetail.EOF = False
      InventoryAcct$ = ""
      COGSAccount$ = ""
      If InventoryAcct$ = "" Then
        If rsDetail("AR SALED Row Type") = "N" Then ' process non-stock
          InventoryAcct$ = rsCompany("SYS COM Sales Misc Acct")
          COGSAccount$ = rsCompany("SYS COM Sales COGS Acct")
        Else
          'rsInventory.Index = "PrimaryKey"
          rsInventory.MoveFirst
          rsInventory.Find "[INV ITEM Id]='" & rsDetail("AR SALED Item ID") & "'"
          If rsInventory.EOF = True Then
            InventoryAcct$ = rsCompany("SYS COM Sales Inventory Acct")
            COGSAccount$ = rsCompany("SYS COM Sales COGS Acct")
          Else
            InventoryAcct$ = IIf(IsNull(rsInventory("INV ITEM Inventory Account")), rsCompany("SYS COM Sales Inventory Acct"), rsInventory("INV ITEM Inventory Account"))
            COGSAccount$ = IIf(IsNull(rsInventory("INV ITEM Cost Of Sales Account")), rsCompany("SYS COM Sales COGS Acct"), rsInventory("INV ITEM Cost Of Sales Account"))
            InventoryCost@ = IIf(IsNull(rsInventory("INV ITEM Average Cost")), 0, rsInventory("INV ITEM Average Cost"))
          End If
        End If
      End If
      InventoryCost@ = 0
      If rsDetail("AR SALED Row Type") = "N" Then ' process non-stock
      Else
        InventoryCost@ = LookRecord("[INV ITEM Average Cost]", "[INV Items]", db, "[INV ITEM ID] = '" & rsDetail("AR SALED Item ID") & "'")
      End If

      If InventoryCost@ = 0 Then InventoryCost@ = IIf(IsNull(rsDetail("AR SALED Item Cost")), 0, rsDetail("AR SALED Item Cost"))

      Dim TempCost#
      TempCost# = InventoryCost@
      InventoryCost@ = Round(TempCost#)

      'Don't do a GetUOMMultiplier on non-stock items
      If rsDetail("AR SALED Row Type") = "N" Then
        QtySold# = rsDetail("AR SALED Qty")
      Else
        QtySold# = rsDetail("AR SALED Qty") * GetUOMMultiplier(rsDetail("AR SALED Item ID"), rsDetail("AR SALED Units"), db)
      End If

      InventoryCost@ = InventoryCost@ * QtySold#

      InventoryCost@ = DropAllBut2(InventoryCost@)

      If InventoryCost@ = 0 Then
      Else
        COGS@ = COGS@ + InventoryCost@
        SQLstatement = "INSERT INTO [GL Work Detail]"
        SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
        SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & InventoryAcct$ & "'," & InventoryCost@ & ",0)"
        db.Execute SQLstatement
        'rsGLWorkDetail.AddNew
        '  rsGLWorkDetail("GW TRANSD Number") = NewNumber&
        '  rsGLWorkDetail("GW TRANSD Account") = InventoryAcct$
        '  rsGLWorkDetail("GW TRANSD Debit Amount") = InventoryCost@
        '  rsGLWorkDetail("GW TRANSD Credit Amount") = 0
        '  rsGLWorkDetail("GW TRANSD Project") = ""
        'rsGLWorkDetail.Update
        
        ' debit for COGS
        SQLstatement = "INSERT INTO [GL Work Detail]"
        SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
        SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & COGSAccount$ & "',0," & InventoryCost@ & ")"
        db.Execute SQLstatement
        'rsGLWorkDetail.AddNew
        '  rsGLWorkDetail("GW TRANSD Number") = NewNumber&
        '  rsGLWorkDetail("GW TRANSD Account") = COGSAccount$
        '  rsGLWorkDetail("GW TRANSD Debit Amount") = 0
         ' rsGLWorkDetail("GW TRANSD Credit Amount") = InventoryCost@
        '  rsGLWorkDetail("GW TRANSD Project") = ""
        'rsGLWorkDetail.Update
                
      End If
      rsDetail.MoveNext
      'If Err = 3021 Then Exit Do
    Loop
    rsInventory.Close
    Set rsInventory = Nothing
  End If
  
  ' post GL entry for this Return
  Dim Success%
  Success% = PostGLWorkDetail(TranDate, NewNumber&, db)
  If Success% = False Then
    MsgBox "An error occurred writing GL Transaction!", , "Error"
    PostReturn = False
    Exit Function
  End If

PostReturns_Exit:

  PostReturn = True
  'rsGLWorkDetail.Close
  'Set rsGLWorkDetail = Nothing
  rsDetail.Close
  Set rsDetail = Nothing
'  rsTaxGroupDetail.Close
'  Set rsTaxGroupDetail = Nothing
  'rsGLTrans.Close
  'Set rsGLTrans = Nothing
  rsCompany.Close
  Set rsCompany = Nothing
  rsSales.Close
  Set rsSales = Nothing
  'rsARPaymentHeader.Close
  'Set rsARPaymentHeader = Nothing
  'rsARCross.Close
  'Set rsARCross = Nothing
  'rsCustomer.Close
  '''Set rsCustomer = Nothing
'  rsCompany.Close
  'St rsCompany = Nothing
  'rsInventory.Close
  'Set rsInventory = Nothing
  'db.Close
  'Set db = Nothing

  Exit Function

PostReturns_Error:
  Call ErrorLog("Sales Module", "PostReturn", Now, Err.Number, Err.Description, intShowError, db)
  PostReturn = False
  'rsGLWorkDetail.Close
  'Set rsGLWorkDetail = Nothing
  rsDetail.Close
  Set rsDetail = Nothing
'  rsTaxGroupDetail.Close
'  Set rsTaxGroupDetail = Nothing
  'rsGLTrans.Close
  'Set rsGLTrans = Nothing
  rsCompany.Close
  Set rsCompany = Nothing
  rsSales.Close
  Set rsSales = Nothing
  'rsARPaymentHeader.Close
  'Set rsARPaymentHeader = Nothing
  'rsARCross.Close
  'Set rsARCross = Nothing
  'rsCustomer.Close
  'Set rsCustomer = Nothing
  'rsCompany.Close
  'Set rsCompany = Nothing
  'rsInventory.Close
  'Set rsInventory = Nothing
  'db.Close
  'Set db = Nothing
  Exit Function


UnableToPostReturnHere:
  PostReturn = False
  'rsGLWorkDetail.Close
  'Set rsGLWorkDetail = Nothing
  rsDetail.Close
  Set rsDetail = Nothing
'  rsTaxGroupDetail.Close
'  Set rsTaxGroupDetail = Nothing
  'rsGLTrans.Close
  'Set rsGLTrans = Nothing
  rsCompany.Close
  Set rsCompany = Nothing
  rsSales.Close
  Set rsSales = Nothing
  'rsARPaymentHeader.Close
  'Set rsARPaymentHeader = Nothing
  'rsARCross.Close
  'Set rsARCross = Nothing
  'rsCustomer.Close
  'Set rsCustomer = Nothing
  'rsCompany.Close
  'Set rsCompany = Nothing
  'rsInventory.Close
  'Set rsInventory = Nothing
  'db.Close
  'Set db = Nothing
  Exit Function

  'rsGLWorkDetail.Close
  'Set rsGLWorkDetail = Nothing
  rsDetail.Close
  Set rsDetail = Nothing
'  rsTaxGroupDetail.Close
'  Set rsTaxGroupDetail = Nothing
  'rsGLTrans.Close
  'Set rsGLTrans = Nothing
  rsCompany.Close
  Set rsCompany = Nothing
  rsSales.Close
  Set rsSales = Nothing
  'rsARPaymentHeader.Close
  'Set rsARPaymentHeader = Nothing
  'rsARCross.Close
  'Set rsARCross = Nothing
  'rsCustomer.Close
  'Set rsCustomer = Nothing
  'rsCompany.Close
  'Set rsCompany = Nothing
  rsInventory.Close
  Set rsInventory = Nothing
  'db.Close
  'Set db = Nothing

End Function

Function PostSalesMemo(DocumentKey&, intShowError As Integer, db As ADODB.Connection)

  ''On Error GoTo PostSalesMemo_error

  Dim msg$
  Dim title$

  'Dim db As ADODB.Connection
  'Set db = New ADODB.Connection
  'db.CursorLocation = adUseServer
  'db.Open gblADOProvider

  Dim rsCompany As ADODB.Recordset
  Set rsCompany = New ADODB.Recordset
  rsCompany.Open "SELECT [SYS COM GL Post By Date],[SYS COM Sales Acct Default]," & _
  "[SYS COM Inventory Cost Method Last YN],[SYS COM Sales AR Acct],[SYS COM Sales Discount Acct]," & _
  "[SYS COM Sales Sales Acct],[SYS COM Sales Misc Acct],[SYS COM Sales Sales Tax]," & _
  "[SYS COM Sales Freight Acct],[SYS COM Sales COGS Acct],[SYS COM Sales Inventory Acct] " & _
  "FROM [SYS Company]", db, adOpenKeyset, adLockOptimistic, adCmdText
  
  rsCompany.MoveFirst
  
  'Dim rsGLWorkDetail As ADODB.Recordset
  'Set rsGLWorkDetail = New ADODB.Recordset
  'rsGLWorkDetail.Open "[GL Work Detail]", db, adOpenStatic, adLockOptimistic, adCmdTable

  Dim rsSales As ADODB.Recordset
  Set rsSales = New ADODB.Recordset
  rsSales.Open "SELECT [AR SALE Date],[AR SALE Customer ID],[AR SALE Total],[AR SALE Ext Document #]," & _
  "[AR SALE Check Acct ID],[AR SALE Billing Customer],[AR SALE Bill To],[AR SALE Description]," & _
  "[AR SALE Discount Amount],[AR SALE Sales Tax],[AR SALE Tax Group],[AR SALE Taxable Subtotal]," & _
  "[AR SALE Freight],[AR SALE Document #],[AR SALE Amount Paid],[AR SALE Document Type]" & _
  "FROM [AR Sales] WHERE [AR SALE Document #]=" & DocumentKey&, db, adOpenKeyset, adLockOptimistic, adCmdText
  
  'rsSales.Index = "PrimaryKey"
 ' rsSales.MoveFirst
  'rsSales.Find "[AR SALE Document #]=" & DocumentKey&

  'Post by 1 - system date or 2 - Transaction date?
  Dim PostDate%
  PostDate% = rsCompany("SYS COM GL Post By Date")

  'Set Invoice Type
  Dim InvoiceType$
  InvoiceType$ = rsSales("AR SALE Document Type")

  'Set Post Date
  Dim TranDate As Variant
  If PostDate% = 1 Then
    TranDate = DateValue(Format(Now, "Short Date"))
  Else
    TranDate = DateValue(rsSales("AR SALE Date"))
  End If
  
  'Verify period can be posted to
  'Send TranDate
  'Return PeriodToPost and PeriodClosed
  Dim PeriodToPost%
  Dim PeriodClosed%
  Call VerifyPeriod(TranDate, PeriodToPost%, PeriodClosed%, db)
  
  'Is period open?
  If PeriodClosed% = True Then
    MsgBox "Unable to post transaction to a closed period.", , "Post Invoice Error"
    GoTo UnableToPostMemoHere
  End If

  ''On Error GoTo PostSalesMemo_error
  
  ' clear any GL Work records
  'Dim cmdtemp As ADODB.Recordset
  'Set cmdtemp = New ADODB.Recordset
  'cmdtemp.Open "DELETE DISTINCTROW * FROM [GL Work Detail]", db, , , adCmdText
  'cmdtemp.Close
  'Set cmdtemp = Nothing
  db.Execute "DELETE DISTINCTROW * FROM [GL Work Detail]", , adCmdText

  Dim BankKey$
  BankKey$ = IIf(IsNull(rsSales("AR SALE Check Acct ID")), "", rsSales("AR SALE Check Acct ID"))

  Dim AppliedAmount@
  If IsNull(rsSales("AR SALE Amount Paid")) Then
    AppliedAmount@ = 0
  Else
    AppliedAmount@ = rsSales("AR SALE Amount Paid")
  End If

  ' determine if we are to process cash or <> cash


  ' update customer stats
  Dim rsCustomer As ADODB.Recordset
  Dim CurrentBalance@
  Set rsCustomer = New ADODB.Recordset
  rsCustomer.Open "SELECT [AR CUST Customer ID],[AR CUST Payments YTD],[AR CUST Payments Lifetime],[AR CUST Sales YTD]," & _
  "[AR CUST Sales Lifetime],[AR CUST Invoices Lifetime],[AR CUST Invoices YTD],[AR CUST Financial Period 1]," & _
  "[AR CUST Highest Balance] FROM [AR Customer] WHERE [AR CUST Customer ID] = '" & rsSales("AR SALE Customer ID") & "'", db, adOpenStatic, adLockOptimistic, adCmdText

  ' if we are to post by customer get the Customer Sales Acct
  Dim SalesAcctDefault%
  Dim CustomerSalesAcct$
  '1-Post using customer accounts; 2-Post using item accounts
  SalesAcctDefault% = rsCompany("SYS COM Sales Acct Default")
  CustomerSalesAcct$ = ""
  If SalesAcctDefault% = 1 Then
    CustomerSalesAcct$ = IIf(IsNull(rsCustomer("AR CUST Sales Account")), "", rsCustomer("AR CUST Sales Account"))
  End If

    rsCustomer("AR CUST Payments YTD") = rsCustomer("AR CUST Payments YTD") + rsSales("AR SALE Amount Paid")
    rsCustomer("AR CUST Payments Lifetime") = rsCustomer("AR CUST Payments Lifetime") + rsSales("AR SALE Amount Paid")
    rsCustomer("AR CUST Sales YTD") = rsCustomer("AR CUST Sales YTD") + rsSales("AR SALE Total")
    rsCustomer("AR CUST Sales Lifetime") = rsCustomer("AR CUST Sales Lifetime") + rsSales("AR SALE Total")
    rsCustomer("AR CUST Invoices Lifetime") = rsCustomer("AR CUST Invoices Lifetime") + 1
    rsCustomer("AR CUST Invoices YTD") = rsCustomer("AR CUST Invoices YTD") + 1

    ' Update current Balance - if not Paid in full
    If rsSales("AR SALE Total") - rsSales("AR SALE Amount Paid") > 0 Then
      CurrentBalance@ = IIf(IsNull(rsCustomer("AR CUST Financial Period 1")), 0, rsCustomer("AR CUST Financial Period 1"))
      CurrentBalance@ = CurrentBalance@ + (rsSales("AR SALE Total") - rsSales("AR SALE Amount Paid"))
      rsCustomer("AR CUST Financial Period 1") = CurrentBalance@
      If CurrentBalance@ > IIf(IsNull(rsCustomer("AR CUST Highest Balance")), 0, rsCustomer("AR CUST Highest Balance")) Then rsCustomer("AR CUST Highest Balance") = CurrentBalance@
    End If
  rsCustomer.Update
  rsCustomer.Close
  Set rsCustomer = Nothing
  '--------------------------------------------------
  ' New AR Cross Payment and AR Payment Header

  Dim PaymentID&

  If AppliedAmount@ > 0 Then
    ' New AR Cross Payment and AR Payment Header
    'Dim rsARPaymentHeader As ADODB.Recordset
    'Set rsARPaymentHeader = New ADODB.Recordset
    'rsARPaymentHeader.Open "[AR Payment Header]", db, adOpenStatic, adLockOptimistic, adCmdTable
    
    'Dim rsARCross As ADODB.Recordset
    'Set rsARCross = New ADODB.Recordset
    'rsARCross.Open "[AR Payment Invoice Cross Reference]", db, adOpenStatic, adLockOptimistic, adCmdTable
    '[INV ITEM Id],[INV ITEM Average Cost],[INV ITEM Standard Cost],[AR SALED Item Cost]
    '[INV ITEM Qty On Hand],[INV ITEM Inventory Account],[INV ITEM Cost Of Sales Account]
    '  write Payment Header
    'rsARPaymentHeader.AddNew
    '  rsARPaymentHeader("AR PAY Type") = "Payment Invoice"
    '  rsARPaymentHeader("AR PAY Check No") = rsSales("AR SALE Check Number") & ""
    '  rsARPaymentHeader("AR PAY Customer No") = rsSales("AR SALE Customer ID") & ""
    '  rsARPaymentHeader("AR PAY Transaction Date") = rsSales("AR SALE Date")
    '  rsARPaymentHeader("AR PAY Amount") = rsSales("AR SALE Amount Paid")
    '  rsARPaymentHeader("AR PAY UnApplied Amount") = 0  'Cannot have unapplied amounts here
    '  rsARPaymentHeader("AR PAY Bank Account") = BankKey$
    '  rsARPaymentHeader("AR PAY Status") = "Posted"
    '  rsARPaymentHeader("AR PAY Posted YN") = True
    '  rsARPaymentHeader("AR PAY NSF") = False
    '  rsARPaymentHeader("AR PAY Cleared") = False
    'rsARPaymentHeader.Update
    'PaymentID& = rsARPaymentHeader("AR PAY ID")
    ' end of write payment header
      SQLstatement = "INSERT INTO [AR Payment Header]"
      SQLstatement = SQLstatement & " ([AR PAY Type],[AR PAY Check No],[AR PAY Customer No],[AR PAY Transaction Date],[AR PAY Amount],[AR PAY UnApplied Amount],[AR PAY Bank Accoun],[AR PAY Status],[AR PAY Posted YN],[AR PAY NSF],[AR PAY Cleared])"
      SQLstatement = SQLstatement & " VALUES ('Credit Memo','" & CStr(rsSales("AR SALE Check Number")) & "','" & CStr(rsSales("AR SALE Customer ID")) & "',#" & rsSales("AR SALE Date") & "#," & rsSales("AR SALE Amount Paid") & ",0,'" & BankKey$ & "','Posted',True,False,False)"
      db.Execute SQLstatement

      Dim rsARPaymentHeader As ADODB.Recordset
      Set rsARPaymentHeader = New ADODB.Recordset
      rsARPaymentHeader.Open "SELECT [AR PAY ID] FROM [AR Payment Header] WHERE [AR PAY Check No]='" & rsSales("AR SALE Check Number") & "'", db, adOpenKeyset, adLockReadOnly, adCmdText
      If rsARPaymentHeader.RecordCount > 1 Then
        rsARPaymentHeader.MoveLast
        PaymentID& = rsARPaymentHeader("AR PAY ID")
      Else
        PaymentID& = rsARPaymentHeader("AR PAY ID")
      End If
      rsARPaymentHeader.Close
      Set rsARPaymentHeader = Nothing
    
    ' write AR Cross reference Record
    'rsARCross.AddNew
    '  rsARCross("AR CROSS Payment ID") = PaymentID&
    '  rsARCross("AR CROSS Payed ID") = rsSales("AR SALE Document #")
    '  rsARCross("AR CROSS Discount Taken") = 0
    '  rsARCross("AR CROSS Write Off Amount") = 0
    '  rsARCross("AR CROSS Applied Amount") = AppliedAmount@
    '  rsARCross("AR CROSS Cleared") = False
    'rsARCross.Update
      SQLstatement = "INSERT INTO [AR Payment Invoice Cross Reference]"
      SQLstatement = SQLstatement & " ([AR CROSS Payment ID],[AR CROSS Payed ID],[AR CROSS Discount Taken],[AR CROSS Write Off Amount],[AR CROSS Applied Amount],[AR CROSS Cleared])"
      SQLstatement = SQLstatement & " VALUES (" & PaymentID& & "," & rsSales("AR SALE Document #") & ",0,0," & AppliedAmount@ & ",False)"
      db.Execute SQLstatement
    
  End If
  ' end of write AR Cross reference Record


  ' End of AR Cross Payment and AR Payment Header
  '--------------------------------------------------


  'Dim rsGLTrans As ADODB.Recordset
  'Set rsGLTrans = New ADODB.Recordset
  'rsGLTrans.Open "[GL Transaction]", db, adOpenStatic, adLockOptimistic, adCmdTable
                    
  ' write GL Transaction Header
  Dim refr$
  Dim desc$
  Dim NewNumber&
  'rsGLTrans.AddNew
      SQLstatement = "INSERT INTO [GL Transaction]"
      SQLstatement = SQLstatement & " ([GL TRANS Document #],[GL TRANS Type],[GL TRANS Date],"
      SQLstatement = SQLstatement & " [GL TRANS Reference],[GL TRANS Amount],[GL TRANS Posted YN],"
      SQLstatement = SQLstatement & " [GL TRANS Description],[GL TRANS Source])"
    
    'rsGLTrans("GL TRANS Document #") = "SALES MEMO " & rsSales("AR SALE Ext Document #")
    'rsGLTrans("GL TRANS Type") = "Sales Memo"
    
    ' gl post date
    If PostDate% = 1 Then
      TempStr = Format(Now, "Short Date")
    Else
      TempStr = rsSales("AR SALE Date")
    End If
    
    SQLstatement = SQLstatement & " VALUES ('SALES MEMO " & rsSales("AR SALE Ext Document #") & "','Sales Memo',#" & TempStr & "#,"
    
    If IsNull(rsSales("AR SALE Billing Customer")) Then
        refr$ = rsSales("AR SALE Bill To") & ""
    Else
        refr$ = rsSales("AR SALE Billing Customer") & ""
    End If
    
    desc$ = IIf(IsNull(rsSales("AR SALE Description")), "", rsSales("AR SALE Description"))
    If Len(Trim$(desc$)) = 0 Then
      desc$ = "SALES MEMO " & rsSales("AR SALE Ext Document #")
    End If
      
      SQLstatement = SQLstatement & "'" & refr$ & "'," & rsSales("AR SALE Total") & ",1,"
      SQLstatement = SQLstatement & "'" & desc$ & "','SALES MEMO " & rsSales("AR SALE Ext Document #") & "',True)"
      'Debug.Print SQLstatement
      
      db.Execute SQLstatement
    
  '  rsGLTrans("GL TRANS Reference") = refr$
  '  rsGLTrans("GL TRANS Amount") = rsSales("AR SALE Total")
  '  rsGLTrans("GL TRANS Posted YN") = 1
  '  rsGLTrans("GL TRANS Description") = desc$
  '  rsGLTrans("GL TRANS Source") = "SALES MEMO " & rsSales("AR SALE Ext Document #")
  '  rsGLTrans("GL TRANS System Generated") = True
  'rsGLTrans.Update
  Dim rsGLTrans As ADODB.Recordset
  Set rsGLTrans = New ADODB.Recordset
  rsGLTrans.Open "SELECT [GL TRANS Number] FROM [GL Transaction] WHERE [GL TRANS Document #]='" & "SALES MEMO " & rsSales("AR SALE Ext Document #") & "'", db, adOpenKeyset, adLockReadOnly, adCmdText
      NewNumber& = rsGLTrans("GL TRANS Number")
  rsGLTrans.Close
  Set rsGLTrans = Nothing
  ' write GL Transaction Detail


  ' update GL
  '  Transaction #1
  '-----------------------------------------------------------------------
  ' Invoice with Payment GL Affected Accounts
  '
  '                  Debit   Credit   Source
  '                  -----   ------   ------
  ' AR                 X              Pref - Sales
  ' CASH               X              Bank - Cash Acct
  ' Discount           X              Pref - Sales
  ' Sales                      X      Item - Inventory or Customer - gSalesAcctDefault%
  ' Sales Tax                  X      Sys Tax
  ' Freight Income             X      Pref - Sales
  ' The following entry is only valid if the payment exceeds the invoice total
  ' AR                         X      Pref - Sales
  '
  '-----------------------------------------------------------------------

  ' Debits
  ' AR

  If rsSales("AR SALE Total") - rsSales("AR SALE Amount Paid") > 0 Then
      SQLstatement = "INSERT INTO [GL Work Detail]"
      SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
      SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & rsCompany("SYS COM Sales AR Acct") & "" & "'," & rsSales("AR SALE Total") - rsSales("AR SALE Amount Paid") & ",0)"
      db.Execute SQLstatement
    'rsGLWorkDetail.AddNew
    '  rsGLWorkDetail("GW TRANSD Number") = NewNumber&
    '  rsGLWorkDetail("GW TRANSD Account") = rsCompany("SYS COM Sales AR Acct")
    '  rsGLWorkDetail("GW TRANSD Debit Amount") = rsSales("AR SALE Total") - rsSales("AR SALE Amount Paid")
    '  rsGLWorkDetail("GW TRANSD Credit Amount") = 0
    '  rsGLWorkDetail("GW TRANSD Project") = ""
    'rsGLWorkDetail.Update
  End If
  
  ' Cash Receipt
  If rsSales("AR SALE Amount Paid") > 0 Then
      SQLstatement = "INSERT INTO [GL Work Detail]"
      SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
      SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & BankKey$ & "'," & rsSales("AR SALE Amount Paid") & ",0)"
      db.Execute SQLstatement
    'rsGLWorkDetail.AddNew
    '  rsGLWorkDetail("GW TRANSD Number") = NewNumber&
    '  rsGLWorkDetail("GW TRANSD Account") = BankKey$
    '  rsGLWorkDetail("GW TRANSD Debit Amount") = rsSales("AR SALE Amount Paid")
    '  rsGLWorkDetail("GW TRANSD Credit Amount") = 0
    '  rsGLWorkDetail("GW TRANSD Project") = ""
    'rsGLWorkDetail.Update
  End If
  
  ' Discount Amount
  If rsSales("AR SALE Discount Amount") > 0 Then
      SQLstatement = "INSERT INTO [GL Work Detail]"
      SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
      SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & rsCompany("SYS COM Sales Discount Acct") & "" & "'," & rsSales("AR SALE Discount Amount") & ",0)"
      db.Execute SQLstatement
    'rsGLWorkDetail.AddNew
    '  rsGLWorkDetail("GW TRANSD Number") = NewNumber&
    '  rsGLWorkDetail("GW TRANSD Account") = rsCompany("SYS COM Sales Discount Acct")
    '  rsGLWorkDetail("GW TRANSD Debit Amount") = rsSales("AR SALE Discount Amount")
    '  rsGLWorkDetail("GW TRANSD Credit Amount") = 0
    '  rsGLWorkDetail("GW TRANSD Project") = ""
    'rsGLWorkDetail.Update
  End If
  ' Sales Tax
  Dim GLTaxAcct$
  Dim TaxGroup$
  Dim TaxPercent#
  Dim TaxID$
  Dim Taxtotal#
  
  If rsSales("AR SALE Sales Tax") > 0 Then
    GLTaxAcct$ = ""
    TaxGroup$ = IIf(IsNull(rsSales("AR SALE Tax Group")), rsCompany("SYS COM Sales Sales Tax"), rsSales("AR SALE Tax Group"))
    
    Dim rsTaxGroupDetail As ADODB.Recordset
    Set rsTaxGroupDetail = New ADODB.Recordset
    rsTaxGroupDetail.Open "SELECT * FROM [SYS Tax Group Detail] where [SYS TAXGRPD Group ID] = '" & TaxGroup$ & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    
    If rsTaxGroupDetail.RecordCount = 0 Then
    Else
      rsTaxGroupDetail.MoveFirst
      Do While rsTaxGroupDetail.EOF = False
      
        TaxPercent# = 0
        TaxID$ = rsTaxGroupDetail("SYS TAXGRPD Tax ID")
        TaxPercent# = LookRecord("[SYS Tax Percent]", "[SYS Tax]", db, "[SYS Tax ID] = '" & TaxID$ & "'")
        GLTaxAcct$ = LookRecord("[SYS Tax Account]", "[SYS Tax]", db, "[SYS Tax ID] = '" & TaxID$ & "'")
        
        If GLTaxAcct$ = "" Then
          GLTaxAcct$ = rsCompany("SYS COM Sales Sales Acct")
        End If
        Taxtotal# = Round(rsSales("AR SALE Taxable Subtotal") * (TaxPercent# / 100))
        SQLstatement = "INSERT INTO [GL Work Detail]"
        SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
        SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & GLTaxAcct$ & "',0," & Taxtotal# & ")"
        db.Execute SQLstatement
      
        'rsGLWorkDetail.AddNew
        '  rsGLWorkDetail("GW TRANSD Number") = NewNumber&
        '  rsGLWorkDetail("GW TRANSD Account") = GLTaxAcct$
        '  rsGLWorkDetail("GW TRANSD Debit Amount") = 0
        '  rsGLWorkDetail("GW TRANSD Credit Amount") = Round(rsSales("AR SALE Taxable Subtotal") * (TaxPercent# / 100))
        '  rsGLWorkDetail("GW TRANSD Project") = ""
        'rsGLWorkDetail.Update

        rsTaxGroupDetail.MoveNext
      Loop
      rsTaxGroupDetail.Close
      Set rsTaxGroupDetail = Nothing
    End If
  End If
  ' Sales Tax
  
  ' Freight Income
  If rsSales("AR SALE Freight") > 0 Then
    SQLstatement = "INSERT INTO [GL Work Detail]"
    SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
    SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & rsCompany("SYS COM Sales Freight Acct") & "" & "',0," & rsSales("AR SALE Freight") & ")"
    db.Execute SQLstatement
    'rsGLWorkDetail.AddNew
    '  rsGLWorkDetail("GW TRANSD Number") = NewNumber&
    '  rsGLWorkDetail("GW TRANSD Account") = rsCompany("SYS COM Sales Freight Acct")
    '  rsGLWorkDetail("GW TRANSD Debit Amount") = 0
    '  rsGLWorkDetail("GW TRANSD Credit Amount") = rsSales("AR SALE Freight")
    '  rsGLWorkDetail("GW TRANSD Project") = ""
    'rsGLWorkDetail.Update
  End If

  ' Credits
  ' Sales Increase
  
  Dim Longer&
  Dim InventoryAcct$
  
  ' item
  Dim rsDetail As ADODB.Recordset
  Set rsDetail = New ADODB.Recordset
  Longer& = 0
  rsDetail.Open "SELECT [AR SALED Item ID],[AR SALED Item Cost],[AR SALED Qty],[AR SALED Item Total],[AR SALED Units],[AR SALED Posting Account],[AR SALED Row Type] FROM [AR Sales Detail] where [AR SALED Document #] = " & rsSales("AR SALE Document #"), db, adOpenKeyset, adLockOptimistic, adCmdText
  
  'rsDetail.MoveLast
  'rsDetail.MoveFirst
  If rsDetail.RecordCount = 0 Then
    ' no detail data
  Else
    rsDetail.MoveFirst
    ' may have detail data
    Do While rsDetail.EOF = False
        SQLstatement = "INSERT INTO [GL Work Detail]"
        SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
        SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & rsDetail("AR SALED Posting Account") & "',0," & rsDetail("AR SALED Item Total") & ")"
        db.Execute SQLstatement
      'rsGLWorkDetail.AddNew
      '  rsGLWorkDetail("GW TRANSD Number") = NewNumber&
      '  rsGLWorkDetail("GW TRANSD Account") = rsDetail("AR SALED Posting Account")
      '  rsGLWorkDetail("GW TRANSD Debit Amount") = 0
      '  rsGLWorkDetail("GW TRANSD Credit Amount") = rsDetail("AR SALED Item Total")
      '  rsGLWorkDetail("GW TRANSD Project") = ""
      'rsGLWorkDetail.Update

      rsDetail.MoveNext
      'If Err = 3021 Then Exit Do
    Loop
  End If
  
      
  ' AR - AR is only credited if the payment amount exceeds the invoice total
  If rsSales("AR SALE Amount Paid") > rsSales("AR SALE Total") Then
    SQLstatement = "INSERT INTO [GL Work Detail]"
    SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
    SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & rsCompany("SYS COM Sales AR Acct") & "',0," & rsSales("AR SALE Amount Paid") - rsSales("AR SALE Total") & ")"
    db.Execute SQLstatement
    'rsGLWorkDetail.AddNew
    '  rsGLWorkDetail("GW TRANSD Number") = NewNumber&
    '  rsGLWorkDetail("GW TRANSD Account") = rsCompany("SYS COM Sales AR Acct")
    '  rsGLWorkDetail("GW TRANSD Debit Amount") = 0
    '  rsGLWorkDetail("GW TRANSD Credit Amount") = rsSales("AR SALE Amount Paid") - rsSales("AR SALE Total")
    '  rsGLWorkDetail("GW TRANSD Project") = ""
    'rsGLWorkDetail.Update
  End If

  ' post GL entry for this Invoice & Payment
  
  Dim Success%
  Success% = PostGLWorkDetail(TranDate, NewNumber&, db)
  If Success% = False Then
    MsgBox "An error occurred writing GL Transaction!", , "Error"
    PostSalesMemo = False
    Exit Function
  End If

PostSalesMemo_Exit:

  PostSalesMemo = True
  'rsGLWorkDetail.Close
  'Set rsGLWorkDetail = Nothing
  rsDetail.Close
  Set rsDetail = Nothing
'  rsTaxGroupDetail.Close
'  Set rsTaxGroupDetail = Nothing
  'rsGLTrans.Close
  'Set rsGLTrans = Nothing
  rsCompany.Close
  Set rsCompany = Nothing
  rsSales.Close
  Set rsSales = Nothing
  'rsARPaymentHeader.Close
  'Set rsARPaymentHeader = Nothing
  'rsARCross.Close
  'Set rsARCross = Nothing
  'rsCustomer.Close
  'Set rsCustomer = Nothing
'  rsCompany.Close
  'Set rsCompany = Nothing
  'db.Close
  'Set db = Nothing
  Exit Function

PostSalesMemo_error:
  Call ErrorLog("Sales Module", "PostSalesMemo", Now, Err.Number, Err.Description, intShowError, db)
  PostSalesMemo = False
  'rsGLWorkDetail.Close
  'Set rsGLWorkDetail = Nothing
  rsDetail.Close
  Set rsDetail = Nothing
  'rsTaxGroupDetail.Close
  'Set rsTaxGroupDetail = Nothing
  'rsGLTrans.Close
  'Set rsGLTrans = Nothing
  rsCompany.Close
  Set rsCompany = Nothing
  rsSales.Close
  Set rsSales = Nothing
  'rsARPaymentHeader.Close
  'Set rsARPaymentHeader = Nothing
  'rsARCross.Close
  'Set rsARCross = Nothing
  'rsCustomer.Close
  'Set rsCustomer = Nothing
  'rsCompany.Close
  'Set rsCompany = Nothing
  'db.Close
  'Set db = Nothing
  Exit Function


UnableToPostMemoHere:
  PostSalesMemo = False
  'rsGLWorkDetail.Close
  'Set rsGLWorkDetail = Nothing
  rsDetail.Close
  Set rsDetail = Nothing
  'rsTaxGroupDetail.Close
  'Set rsTaxGroupDetail = Nothing
  'rsGLTrans.Close
  'Set rsGLTrans = Nothing
  rsCompany.Close
  Set rsCompany = Nothing
  rsSales.Close
  Set rsSales = Nothing
  'rsARPaymentHeader.Close
  'Set rsARPaymentHeader = Nothing
  'rsARCross.Close
  'Set rsARCross = Nothing
  'rsCustomer.Close
  'Set rsCustomer = Nothing
  'rsCompany.Close
  'Set rsCompany = Nothing
  'db.Close
  'Set db = Nothing
  Exit Function
  
  'rsGLWorkDetail.Close
  'Set rsGLWorkDetail = Nothing
  rsDetail.Close
  Set rsDetail = Nothing
  'rsTaxGroupDetail.Close
  'Set rsTaxGroupDetail = Nothing
  'rsGLTrans.Close
  'Set rsGLTrans = Nothing
  rsCompany.Close
  Set rsCompany = Nothing
  rsSales.Close
  Set rsSales = Nothing
  'rsARPaymentHeader.Close
  'Set rsARPaymentHeader = Nothing
  'rsARCross.Close
  'Set rsARCross = Nothing
  'rsCustomer.Close
  'Set rsCustomer = Nothing
  'rsCompany.Close
  'Set rsCompany = Nothing
  'db.Close
  'Set db = Nothing

End Function

Function RecurSales()

'  Dim db As ADODB.Connection
'  Set db = New ADODB.Connection
'  db.CursorLocation = adUseServer
'  db.Open gblADOProvider
  
'  Dim rsRecur As ADODB.Recordset
'  Set rsRecur = New ADODB.Recordset
'  rsRecur.Open "SELECT * FROM [AR Sales] where [AR SALE Recur Type] <> 'Never' AND [AR SALE Next Recur] BETWEEN #" & Forms![Recur Sales].[StartDate] & "# AND #" & Forms![Recur Sales].[EndDate] & "#", db, adOpenStatic, adLockOptimistic, adCmdText

'  Dim DocumentKey&
'  Dim Success%
  ''On Error Resume Next
'  rsRecur.MoveFirst
'  If rsRecur.RecordCount = 0 Then GoTo SkipRecurSales

'  Do While Not rsRecur.EOF
'    DocumentKey& = rsRecur("AR SALE Document #")
'    Success% = CloneSales(DocumentKey&, False)
    'Update invoice next recur date
'     Select Case rsRecur("AR SALE Recur Type")
'      Case "Monthly"
'        rsRecur("AR SALE Next Recur") = DateAdd("m", 1, rsRecur("AR SALE Next Recur"))
'      Case "Quarterly"
'        rsRecur("AR SALE Next Recur") = DateAdd("q", 1, rsRecur("AR SALE Next Recur"))
'      Case "Annually"
'        rsRecur("AR SALE Next Recur") = DateAdd("yyyy", 1, rsRecur("AR SALE Next Recur"))
'      End Select
'    rsRecur.Update
'    rsRecur.MoveNext
'  Loop

'SkipRecurSales:
'  rsRecur.Close
'  Set rsRecur = Nothing
'  db.Close
'  Set db = Nothing
'  Exit Function

'  rsRecur.Close
'  Set rsRecur = Nothing
'  db.Close
'  Set db = Nothing
'  Exit Function


End Function

Public Sub AgeSelectedReceivables(CurrentCustomer As String, db As ADODB.Connection)
  
  'On Error GoTo AgeSelectedReceivables_Error

  'Dim db As ADODB.Connection
  'Set db = New ADODB.Connection
  'db.CursorLocation = adUseClient
  'db.Open gblADOProvider
  
  Dim rstInfo As ADODB.Recordset
  Set rstInfo = New ADODB.Recordset
  rstInfo.Open "[exp qryGetPeriod-n-AgingInfo]", db, adOpenStatic, adLockOptimistic, admcdtable
  
  intPeriod1 = rstInfo("SYS COM Sales Period 1")
  intPeriod2 = rstInfo("SYS COM Sales Period 2")
  intPeriod3 = rstInfo("SYS COM Sales Period 3")
  
  'Dim cmdtemp As ADODB.Recordset
  db.Execute "Delete * From [Print Aged Receivables Work]", , adCmdText
  'cmdtemp.Close
  'Set cmdtemp = Nothing
  
  'Pull Discounts Applied to Aged Sales
  Dim qryProcessSelectedARRDiscounts As String
  qryProcessSelectedARRDiscounts = "INSERT INTO [Print Aged Receivables Work] ( [Customer ID], [Transaction Type], [Transaction Date], Balance, [Transaction ID],"
  qryProcessSelectedARRDiscounts = qryProcessSelectedARRDiscounts & " [Transaction Description], [Applied To], Period, [Order] ) SELECT DISTINCTROW [qryCustomerPayments].[AR PAY Customer No] AS [Customer ID], 'Discount' AS [Transaction Type],"
  qryProcessSelectedARRDiscounts = qryProcessSelectedARRDiscounts & " [qryCustomerPayments].[AR PAY Transaction Date] AS [Transaction Date], [AR CROSS Discount Taken]*-1 AS Balance, [qryCustomerPayments].[AR PAY Check No] AS [Transaction ID], 'Applied to ' &"
  qryProcessSelectedARRDiscounts = qryProcessSelectedARRDiscounts & " [AR SALE Ext Document #] AS [Transaction Description], [exp - qryAge Sales].[AR SALE Ext Document #] AS [Applied To], [exp - qryAge Sales].Period AS Expr1, [exp - qryAge Sales].[AR SALE Ext Document #]"
  qryProcessSelectedARRDiscounts = qryProcessSelectedARRDiscounts & " AS Expr2 FROM qryCustomerPayments INNER JOIN [exp - qryAge Sales] ON [qryCustomerPayments].[AR CROSS Payed ID] = [exp - qryAge Sales].[AR SALE Document #] WHERE ((([qryCustomerPayments].[AR PAY Customer No])=[" & CurrentCustomer & "]) AND (([qryCustomerPayments].[AR CROSS Discount Taken])>=0.01))"
  'Dim cmdtemp1 As ADODB.Recordset
  'Debug.Print qryProcessSelectedARRDiscounts
  db.Execute qryProcessSelectedARRDiscounts, , adCmdText
  'cmdtemp1.Close
  'Set cmdtemp1 = Nothing
  
  'Pull Payments on Aged Sales
  Dim qryProcessSelectedARRPayments As String
  qryProcessSelectedARRPayments = qryProcessSelectedARRPayments & "INSERT INTO [Print Aged Receivables Work] ( [Customer ID], [Transaction Type], [Transaction Date], Balance, [Transaction ID], [Transaction Description], [Applied To], Period, [Order] )"
  qryProcessSelectedARRPayments = qryProcessSelectedARRPayments & " SELECT DISTINCTROW [qryCustomerPayments].[AR PAY Customer No] AS [Customer ID], [qryCustomerPayments].[AR PAY Type] AS [Transaction Type], [qryCustomerPayments].[AR PAY Transaction Date]"
  qryProcessSelectedARRPayments = qryProcessSelectedARRPayments & " AS [Transaction Date], [AR CROSS Applied Amount]*-1 AS Balance, [qryCustomerPayments].[AR PAY Check No] AS [Transaction ID], 'Applied to ' & [AR SALE Ext Document #] AS"
  qryProcessSelectedARRPayments = qryProcessSelectedARRPayments & " [Transaction Description], [exp - qryAge Sales].[AR SALE Ext Document #] AS [Applied To], [exp - qryAge Sales].Period AS Expr1, [exp - qryAge Sales].[AR SALE Ext Document #] AS Expr2"
  qryProcessSelectedARRPayments = qryProcessSelectedARRPayments & " FROM [exp - qryAge Sales] INNER JOIN qryCustomerPayments ON [exp - qryAge Sales].[AR SALE Document #] = [qryCustomerPayments].[AR CROSS Payed ID] WHERE ((([qryCustomerPayments].[AR PAY Customer No])=[" & CurrentCustomer & "]) AND (([qryCustomerPayments].[AR CROSS Applied Amount])>=0.01))"
  'Dim cmdtemp2 As ADODB.Recordset
  db.Execute qryProcessSelectedARRPayments, , adCmdText
  'cmdtemp2.Close
  'Set cmdtemp2 = Nothing
  
  'Pull Write Offs on Aged Sales
  Dim qryProcessSelectedARRWriteOffs As String
  qryProcessSelectedARRWriteOffs = "INSERT INTO [Print Aged Receivables Work] ( [Customer ID], [Transaction Type], [Transaction Date], Balance,"
  qryProcessSelectedARRWriteOffs = qryProcessSelectedARRWriteOffs & " [Transaction Description], [Applied To], Period, [Order] ) SELECT DISTINCTROW [qryCustomerPayments].[AR PAY Customer No], 'Write Off' AS [Transaction Type], [qryCustomerPayments].[AR PAY Transaction Date] AS"
  qryProcessSelectedARRWriteOffs = qryProcessSelectedARRWriteOffs & " [Transaction Date], [AR CROSS Write Off Amount]*-1 AS Balance, 'Applied to ' & [AR SALE Ext Document #] AS [Transaction Description], [exp - qryAge Sales].[AR SALE Ext Document #] AS [Applied To], [exp - qryAge Sales].Period AS"
  qryProcessSelectedARRWriteOffs = qryProcessSelectedARRWriteOffs & " Expr1, [exp - qryAge Sales].[AR SALE Ext Document #] AS Expr2 FROM [exp - qryAge Sales] INNER JOIN qryCustomerPayments ON [exp - qryAge Sales].[AR SALE Document #] = [qryCustomerPayments].[AR CROSS Payed ID]WHERE ((([qryCustomerPayments].[AR PAY Customer No])=[" & CurrentCustomer & "]) AND"
  qryProcessSelectedARRWriteOffs = qryProcessSelectedARRWriteOffs & " (([AR CROSS Write Off Amount]*-1)<>0))"
  'Dim cmdtemp3 As ADODB.Recordset
  db.Execute qryProcessSelectedARRWriteOffs, , adCmdText
  'cmdtemp3.Close
  'Set cmdtemp3 = Nothing
  
  'Pull Unapplied Payments
  Dim qryAppendSelectedUnappliedAR As String
  qryAppendSelectedUnappliedAR = "INSERT INTO [Print Aged Receivables Work] ( [Transaction Date], [Customer ID], Balance, Period, [Transaction Description], [Order], [Transaction Type], [Transaction ID] ) SELECT DISTINCTROW [exp - qry Unapplied Payments].[AR PAY Transaction Date], [exp - qry Unapplied Payments].[AR PAY Customer No],"
  qryAppendSelectedUnappliedAR = qryAppendSelectedUnappliedAR & " [AR PAY Unapplied Amount]*-1 AS Amount, 1 AS Period, 'Unapplied' AS [Desc], -1 AS [Order], [exp - qry Unapplied Payments].[AR PAY Type], [exp - qry Unapplied Payments].[AR PAY Check No] FROM [exp - qry Unapplied Payments] WHERE ((([exp - qry Unapplied Payments].[AR PAY Customer No])=[" & CurrentCustomer & "]))"
  'Dim cmdtemp4 As ADODB.Recordset
  db.Execute qryAppendSelectedUnappliedAR, , adCmdText
  'cmdtemp4.Close
  'Set cmdtemp4 = Nothing
  
  'Now add the actual sales
  Dim qryAddSelectedAgedSales As String
  qryAddSelectedAgedSales = "INSERT INTO [Print Aged Receivables Work] ( [Customer ID], Period, [Transaction ID], [Transaction Type], Balance, [Transaction Date], [Order] ) SELECT [exp - qryAge Sales].[AR SALE Customer ID] AS Expr1, [exp - qryAge Sales].Period AS Expr2,"
  qryAddSelectedAgedSales = qryAddSelectedAgedSales & " [exp - qryAge Sales].[AR SALE Ext Document #] AS Expr3, [exp - qryAge Sales].[AR SALE Document Type] AS Expr4, [exp - qryAge Sales].[Transaction Amount] AS Expr5, [exp - qryAge Sales].[AR SALE Date] AS Expr6, [exp - qryAge Sales].[AR SALE Document #] AS Expr7 FROM [exp - qryAge Sales] WHERE ((([exp - qryAge Sales].[AR SALE Customer ID])=[" & CurrentCustomer & "]))"
  'Dim cmdtemp5 As ADODB.Recordset
  db.Execute qryAddSelectedAgedSales, , adCmdText
  'cmdtemp5.Close
  'Set cmdtemp5 = Nothing

  rstInfo.Close
  Set rstInfo = Nothing
  'db.Close
  'Set db = Nothing
  Exit Sub
  
AgeSelectedReceivables_Error:
  Call ErrorLog("Sales Module", "AgeSelectedReceivables", Now, Err.Number, Err.Description, True, db)
  Resume Next
  
  rstInfo.Close
  Set rstInfo = Nothing
  'db.Close
  'Set db = Nothing
End Sub


Function GetSalesAgingPeriod(intDays As Long) As Integer
'Uses Global integer vars holding Aging values intPeriodx
Select Case intDays
Case Is < 0
  GetSalesAgingPeriod = 1
Case 0 To intPeriod1
  GetSalesAgingPeriod = 1
Case intPeriod1 To intPeriod2
  GetSalesAgingPeriod = 2
Case intPeriod2 To intPeriod3
  GetSalesAgingPeriod = 3
Case Else
  GetSalesAgingPeriod = 4
End Select
End Function

Public Sub AgingDetails(CustID As String, db As ADODB.Connection)
ShowStatus True

Dim Period1 As Integer, Period2 As Integer, Period3 As Integer, Period4 As Integer
Dim AgeBy As Integer, Days As Integer
Dim TransDate As Date
Dim SQLstatement As String
    
  Dim rsCompany As ADODB.Recordset
  Set rsCompany = New ADODB.Recordset
  rsCompany.Open "SELECT [SYS COM Sales Age Invoices By],[SYS COM Sales Period 1]," & _
  "[SYS COM Sales Period 2],[SYS COM Sales Period 3] FROM [SYS Company]", db, adOpenKeyset, adLockOptimistic, adCmdText

  rsCompany.MoveFirst
  Period1 = rsCompany("SYS COM Sales Period 1")
  Period2 = rsCompany("SYS COM Sales Period 2")
  Period3 = rsCompany("SYS COM Sales Period 3")
  Period4 = 90
  AgeBy = IIf(IsNull(rsCompany("SYS COM Sales Age Invoices By")), 1, rsCompany("SYS COM Sales Age Invoices By"))
  '1 - Invoice Date  2 - Due Date

  rsCompany.Close
  Set rsCompany = Nothing
  
  db.Execute "DELETE * FROM [AGE Aging Sales Work] WHERE [AGE Cust ID]='" & CustID & "'", , adCmdText
  
  Set ADOsales = New ADODB.Recordset
  ADOsales.Open "SELECT [AR SALE Document Type],[AR SALE Billing Customer],[AR SALE Customer ID],[AR SALE Total]," & _
  "[AR SALE Balance Due],[AR SALE Ext Document #],[AR SALE Date],[AR SALE Due Date] " & _
  "FROM [AR Sales] where [AR SALE Customer ID] = '" & CustID & "' AND [AR SALE Balance Due] > 0 AND [AR SALE Posted YN] = TRUE", db, adOpenKeyset, adLockOptimistic, adCmdText
  'On Error Resume Next
  With ADOsales
    If .RecordCount > 0 Then
      .MoveFirst
      'If Err = 0 Then
      Do While Not ADOsales.EOF
        If AgeBy = 1 Then
            TransDate = ![AR SALE Date]
        Else
            TransDate = ![AR SALE Due Date]
        End If
        Days = DateDiff("d", TransDate, FormatDate(Now))
      
        SQLstatement = "INSERT INTO [AGE Aging Sales Work]"
        SQLstatement = SQLstatement & " ([AGE Cust ID],[AGE Cust Name]," & _
        "[AGE Sales Doc Ext No],[AGE Start Date],[AGE Orig Amount],[AGE Period 1]," & _
        "[AGE Period 2],[AGE Period 3],[AGE Period 4])"
        
        SQLstatement = SQLstatement & " VALUES ('" & ![AR SALE Customer ID] & "','" & _
        ![AR SALE Billing Customer] & "','" & ![AR SALE Ext Document #] & "',#" & _
        TransDate & "#," & ![AR SALE Total] & ","
        
        Select Case Days
        Case Is < 0
                SQLstatement = SQLstatement & ![AR SALE Balance Due] & ",0,0,0" & ")"
                db.Execute SQLstatement, , adCmdText
        Case 0 To Period1
                SQLstatement = SQLstatement & ![AR SALE Balance Due] & ",0,0,0" & ")"
                db.Execute SQLstatement, , adCmdText
       Case Period1 To Period2
                SQLstatement = SQLstatement & "0," & ![AR SALE Balance Due] & ",0,0" & ")"
                db.Execute SQLstatement, , adCmdText
        Case Period2 To Period3
                SQLstatement = SQLstatement & "0,0," & ![AR SALE Balance Due] & ",0" & ")"
                db.Execute SQLstatement, , adCmdText
        Case Else
                SQLstatement = SQLstatement & "0,0,0," & ![AR SALE Balance Due] & ")"
                db.Execute SQLstatement, , adCmdText
        End Select
        
        ADOsales.MoveNext
      Loop
    End If
  End With
    ADOsales.Close
    Set ADOsales = Nothing
ShowStatus False
End Sub

