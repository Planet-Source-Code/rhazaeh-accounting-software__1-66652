Attribute VB_Name = "Year_End_Module"

Function ComputeNetProfit@(Bucket$, ThroughPeriod%, db As ADODB.Connection)
ShowStatus True
  'On Error Resume Next

  'Compute the current net profit from P&L account types
  ' P&L Account Types are Sales, COGS and Expenses

  'Get a dynaset of these types and loop through adding up balances
  
  Dim rs As ADODB.Recordset
  Dim sql$
  Dim Balance@
  Dim Bal@
  Dim X%
  'Dim db As ADODB.Connection
  'Set db = New ADODB.Connection
  'db.CursorLocation = adUseClient
  'db.Open gblADOProvider

  sql$ = "SELECT * FROM [GL Chart of Accounts] "
  sql$ = sql$ & "WHERE [GL COA Asset Type] = 'Sales' OR "
  sql$ = sql$ & "[GL COA Asset Type] = 'Cost of Goods Sold' OR "
  sql$ = sql$ & "[GL COA Asset Type] = 'Expense'"

  Set rs = New ADODB.Recordset
  rs.Open sql$, db, adOpenStatic, adLockOptimistic, adCmdText
  
  'On Error Resume Next
  rs.MoveFirst
  Do While Not rs.EOF
    GoSub subBalance
      Balance@ = Balance@ + Bal@
    rs.MoveNext
  Loop

  ComputeNetProfit@ = Balance@ * -1
  Exit Function

subBalance:
  Bal@ = 0
  Select Case Bucket$
  Case "CY"
    Bal@ = IIf(IsNull(rs("GL COA CY Beginning Amt")), 0, rs("GL COA CY Beginning Amt"))
    For X% = 1 To ThroughPeriod%
      Bal@ = Bal@ + IIf(IsNull(rs("GL COA CY Period " & Trim(Str(X%)) & " Amt")), 0, rs("GL COA CY Period " & Trim(Str(X%)) & " Amt"))
    Next X%
  Case "PY"
    Bal@ = IIf(IsNull(rs("GL COA PY Beginning Amt")), 0, rs("GL COA PY Beginning Amt"))
    For X% = 1 To ThroughPeriod%
      Bal@ = Bal@ + IIf(IsNull(rs("GL COA PY Period " & Trim(Str(X%)) & " Amt")), 0, rs("GL COA PY Period " & Trim(Str(X%)) & " Amt"))
    Next X%
  Case "Budget"
    Bal@ = IIf(IsNull(rs("GL COA BUD Beginning Amt")), 0, rs("GL COA BUD Beginning Amt"))
    For X% = 1 To ThroughPeriod%
      Bal@ = Bal@ + IIf(IsNull(rs("GL COA BUD Period " & Trim(Str(X%)) & " Amt")), 0, rs("GL COA BUD Period " & Trim(Str(X%)) & " Amt"))
    Next X%
  End Select
Return

rs.Close
Set rs = Nothing
'db.Close
'Set db = Nothing
ShowStatus False

End Function

Sub CustomerRollOver(CustomerID$, RollDate As Variant, db As ADODB.Connection)
ShowStatus True
  'On Error Resume Next

  'Dim db As ADODB.Connection
  'Set db = New ADODB.Connection
  'db.CursorLocation = adUseClient
  'db.Open gblADOProvider
  Dim sql$
  Dim rs As ADODB.Recordset
  Dim rs2 As ADODB.Recordset
  Dim rs3 As ADODB.Recordset
  
  'Do a year end rollover up to the given date
  'Need to roll over:
  '   Sales $
  '   Payments $
  '   Write Offs $
  '   # Invoices

  'Find customer record

  Dim rsCustomer As ADODB.Recordset
  Set rsCustomer = New ADODB.Recordset
  rsCustomer.Open "SELECT [AR CUST Customer ID],[AR CUST Sales YTD],[AR CUST Sales Last Year],[AR CUST Invoices YTD]," & _
  "[AR CUST Invoices Last Year],[AR CUST Payments YTD],[AR CUST Payments Last Year]," & _
  "[AR CUST Write Offs YTD],[AR CUST Write Offs Last Year] FROM [AR Customer] " & _
  "WHERE [AR CUST Customer ID]='" & CustomerID$ & "'", db, adOpenKeyset, adLockOptimistic, adCmdText

  'rsCustomer.Index = "PrimaryKey"
  'rsCustomer.Seek CustomerID$
  If rsCustomer.RecordCount = 0 Then
    Exit Sub
  End If

  '----------------------------------------------------------------------
  'Do Sales and # Invoices first
  Dim TempYTD#
  Dim SalesYTD#
  Dim InvoiceCount%
  Dim TempInvoiceCount%
  
  Dim DayOne As Variant
  DayOne = LookRecord("[SYS COM Fiscal Start Date]", "[SYS Company]", db)
  InvoiceCount% = CountRecord("[AR SALE Total]", "[AR Sales]", db, "[AR SALE Customer ID] = '" & CustomerID$ & "' AND [AR SALE Posted YN] = TRUE AND [AR SALE Date] >= #" & DayOne & "# AND [AR SALE Document Type] in ('Invoice','Sales Memo','Beginning Balance','Finance Charge')")
  TempInvoiceCount% = InvoiceCount%

  Dim Sales#
  Dim Returns#
  Sales# = IIf(IsNull(SumRecord("[AR SALE Total]", "[AR Sales]", db, "[AR SALE Customer ID] = '" & CustomerID$ & "' AND [AR SALE Posted YN] = TRUE AND [AR SALE Date] >= #" & DayOne & "# AND [AR SALE Document Type] in ('Invoice','Sales Memo','Beginning Balance','Finance Charge')")), 0, SumRecord("[AR SALE Total]", "AR Sales", db, "[AR SALE Customer ID] = '" & gCustomerID$ & "' AND [AR SALE Posted YN] = TRUE AND [AR SALE Date] >= #" & DayOne & "# AND [AR SALE Document Type] in ('Invoice','Sales Memo','Beginning Balance','Finance Charge')"))
  Returns# = IIf(IsNull(SumRecord("[AR SALE Total]", "[AR Sales]", db, "[AR SALE Customer ID] = '" & CustomerID$ & "' AND [AR SALE Posted YN] = TRUE AND [AR SALE Date] >= #" & DayOne & "# AND [AR SALE Document Type] in ('Return','Credit Memo')")), 0, SumRecord("[AR SALE Total]", "[AR Sales]", db, "[AR SALE Customer ID] = '" & gCustomerID$ & "' AND [AR SALE Posted YN] = TRUE AND [AR SALE Date] >= #" & DayOne & "# AND [AR SALE Document Type] in ('Return','Credit Memo')"))
  SalesYTD# = Sales# - Returns#
  TempYTD# = SalesYTD#
  
  Set rs = New ADODB.Recordset
  rs.Open "SELECT [AR SALE Total] FROM [AR SALES] WHERE [AR SALE Customer ID] = '" & CustomerID$ & "' AND [AR SALE Date] > #" & RollDate & "#", db, adOpenStatic, adLockOptimistic, adCmdText
  'On Error Resume Next
  
  'Err = 0
  If rs.RecordCount > 0 Then
    'Process these sales
    rs.MoveFirst
    Do While Not rs.EOF
      TempYTD# = TempYTD# - rs("AR SALE Total")
      TempInvoiceCount% = TempInvoiceCount% - 1
      rs.MoveNext
    Loop
  'Else
  End If
  rs.Close
  Set rs = Nothing
  
  'Write out new values to customer file
    rsCustomer("AR CUST Sales YTD") = SalesYTD# - TempYTD#
    rsCustomer("AR CUST Sales Last Year") = TempYTD#
    rsCustomer("AR CUST Invoices YTD") = InvoiceCount% - TempInvoiceCount%
    rsCustomer("AR CUST Invoices Last Year") = TempInvoiceCount%
  rsCustomer.Update

  '-----------------------------------------------------------------------
  'Now do payments
  Dim TempPaymentsYTD#
  Dim PaymentsYTD#

  PaymentsYTD# = SumRecord("[AR PAY Amount]", "[AR Payment Header]", db, "[AR PAY Customer No] = '" & CustomerID$ & "' AND [AR PAY Posted YN] = TRUE AND [AR PAY Transaction Date] >= #" & DayOne & "# AND [AR PAY NSF] = FALSE AND [AR PAY Type] <> 'Credit Memo' AND [AR PAY Type] <> 'Return'")
  TempPaymentsYTD# = PaymentsYTD#
    
  Set rs2 = New ADODB.Recordset
  rs2.Open "SELECT [AR PAY Amount] FROM [AR Payment Header] WHERE [AR PAY Customer No] = '" & CustomerID$ & "' AND [AR PAY Transaction Date] > #" & RollDate & "#", db, adOpenStatic, adLockOptimistic, adCmdText
  'On Error Resume Next

  'Err = 0
  'rs2.MoveFirst
  If rs2.RecordCount > 0 Then
    rs2.MoveFirst
    'Process these sales
    Do While Not rs2.EOF
      TempPaymentsYTD# = TempPaymentsYTD# - rs2("AR PAY Amount")
      rs2.MoveNext
    Loop
  End If
  rs2.Close
  Set rs2 = Nothing
  
  'Write out new values to customer file
    rsCustomer("AR CUST Payments YTD") = PaymentsYTD# - TempPaymentsYTD#
    rsCustomer("AR CUST Payments Last Year") = TempPaymentsYTD#
  rsCustomer.Update

  '------------------------------------------------------------------------
  'Now do write offs
  Dim TempWriteOffsYTD#
  Dim WriteOffsYTD#

  WriteOffsYTD# = SumRecord("[AR CROSS Write Off Amount]", "[qryCustomerPayments]", db, "[AR PAY Customer No] = '" & CustomerID$ & "' AND [AR PAY Posted YN] = TRUE AND [AR PAY Transaction Date] >= #" & DayOne & "# AND [AR PAY NSF] = FALSE")
  TempWriteOffsYTD# = WriteOffsYTD#
'again:
  'Check all cross refs to this payment for a writeoff amount
  sql$ = "SELECT [AR CROSS Write Off Amount] FROM [qryCustomerPayments] "
  sql$ = sql$ & " WHERE [AR PAY Customer No] = '" & CustomerID$ & "'"
  sql$ = sql$ & " AND [AR PAY Transaction Date] > #" & RollDate & "#"
  sql$ = sql$ & " AND [AR CROSS Write Off Amount] > 0"
  
  Set rs3 = New ADODB.Recordset
  'GoTo again
  rs3.Open sql$, db, adOpenStatic, adLockOptimistic, adCmdText
  'On Error Resume Next
  
  'Err = 0
  If rs3.RecordCount > 0 Then
  rs3.MoveFirst
    Do While Not rs3.EOF
      TempWriteOffsYTD# = TempWriteOffsYTD# - rs3("AR CROSS Write Off Amount")
      rs3.MoveNext
    Loop
  End If
  'Write out new values to customer file
    rsCustomer("AR CUST Write Offs YTD") = WriteOffsYTD# - TempWriteOffsYTD#
    rsCustomer("AR CUST Write Offs Last Year") = TempWriteOffsYTD#
    rsCustomer.Update

'rs.Close
'Set rs = Nothing
'rs2.Close
'Set rs2 = Nothing
'rs3.Close
'Set rs3 = Nothing
rsCustomer.Close
Set rsCustomer = Nothing
'db.Close
'Set db = Nothing
ShowStatus False

End Sub

Sub DeleteGLTrans(db As ADODB.Connection)
ShowStatus True
  'On Error Resume Next

  'Delete all GL Transactions and detail as of gLastDayOfYear
  'Dim db As ADODB.Connection
  'Set db = New ADODB.Connection
  'db.CursorLocation = adUseClient
  'db.Open gblADOProvider
  
  Dim rs As ADODB.Recordset
  Dim TransNumber&
  Dim sql$
  sql$ = "SELECT * FROM [GL Transaction] where [GL Trans Date] <= #" & gLastDayOfYear & "#"
  sql$ = sql$ & " AND [GL Trans Type] <> 'APT'"

  Set rs = New ADODB.Recordset
  rs.Open sql$, db, adOpenStatic, adLockOptimistic, adCmdText
  'On Error Resume Next

  If rs.RecordCount > 0 Then
  rs.MoveFirst
    Do While Not rs.EOF
      TransNumber& = rs("GL TRANS Number")
      'Find all Detail with this number
      'DeleteGLTransDetail (TransNumber&)
      db.Execute "DELETE * FROM [GL Transaction Detail] where [GL TRANSD Number] = " & TransNumber, , adCmdText
      rs.Delete
      'Err = 0
      rs.MoveNext
      'If Err = 3021 Then Exit Do
    Loop
  End If
ShowStatus False

End Sub

'Sub DeleteGLTransDetail(TransNumber As String)
'Dim db As ADODB.Connection
'Set db = New ADODB.Connection
'db.CursorLocation = adUseClient
'db.Open gblADOProvider
'Dim rs As ADODB.Recordset

'rs.Open "DELETE * FROM [GL Transaction Detail] where [GL TRANSD Number] = " & TransNumber, db, adOpenStatic, adLockOptimistic, adCmdText

'rs.Close
'Set rs = Nothing
'db.Close
'Set db = Nothing

'End Sub


Function RecalcChartOfAccounts(db As ADODB.Connection)
ShowStatus True
  'On Error GoTo RecalcChartOfAccounts_Error
  
  
  Dim Response%
  Response% = MsgBox("Are you sure you want to recalculate the Chart Of Accounts?", vbYesNo + vbQuestion, "Information")
  If Response% = vbNo Then Exit Function
  
  Call ResetCYGLBalances(db)
  
  MsgBox "Recalculation complete."
ShowStatus False
  Exit Function
RecalcChartOfAccounts_Error:
  Call ErrorLog("Year End Module", "RecalcChartOfAccounts", Now, Err.Number, Err.Description, True, db)
  Resume Next
  
End Function

Sub ReloadGLBalances(db As ADODB.Connection)
ShowStatus True
  'On Error Resume Next
  
  'Dim db As ADODB.Connection
  'Set db = New ADODB.Connection
  'db.CursorLocation = adUseClient
  'db.Open gblADOProvider
  
  Dim rsGLTrans As ADODB.Recordset
  rsGLTrans.Open "SELECT * FROM [GL Transaction]", db, adOpenKeyset, adLockOptimistic, adCmdText

  Dim rsGLTransDetail As ADODB.Recordset

  'rsGLTrans.Index = "PrimaryKey"
  'rsGLTransDetail.Index = "GL TRANSD Number"


  Dim Success%
  Dim AccountPost$
  Dim TranDate As Variant
  Dim DebitAmount@
  Dim CreditAmount@

  If rsGLTrans.RecordCount > 0 Then
    rsGLTrans.MoveFirst
    Do While rsGLTrans.EOF = False
    
      If rsGLTrans("GL TRANS Posted YN") = True Then
          
          'rsGLTransDetail.Find rsGLTrans("GL TRANS Number")
          Set rsGLTransDetail = New ADODB.Recordset
          rsGLTransDetail.Open "SELECT [GL TRANSD Number],[GL TRANSD Account]," & _
          "[GL TRANS Date],[GL TRANSD Debit Amount],[GL TRANSD Credit Amount]" & _
          " FROM [GL Transaction Detail] WHERE [GL TRANSD Number]=" & rsGLTrans![GL TRANS Number], db, adOpenKeyset, adLockOptimistic, adCmdText
          rsGLTransDetail.MoveFirst
          
          If rsGLTransDetail.RecordCount > 0 Then
            Do While rsGLTransDetail.EOF = False
              'If rsGLTransDetail("GL TRANSD Number") <> rsGLTrans("GL TRANS Number") Then Exit Do
              AccountPost$ = rsGLTransDetail("GL TRANSD Account")
              TranDate = rsGLTrans("GL TRANS Date")
              DebitAmount@ = rsGLTransDetail("GL TRANSD Debit Amount")
              CreditAmount@ = rsGLTransDetail("GL TRANSD Credit Amount")
              Success% = PostCOA(AccountPost$, TranDate, DebitAmount@, CreditAmount@, db)
              rsGLTransDetail.MoveNext
            Loop
          End If
          rsGLTransDetail.Close
          Set rsGLTransDetail = Nothing
            
      End If
      
      rsGLTrans.MoveNext
    Loop
  End If

  rsGLTrans.Close
  Set rsGLTrans = Nothing
  'rsGLTransDetail.Close
  'Set rsGLTransDetail = Nothing
  'db.Close
  'Set db = Nothing
ShowStatus False
  
End Sub

Sub ResetCYGLBalances(db As ADODB.Connection)
ShowStatus True
  'On Error Resume Next

  'Dim db As ADODB.Connection
  'Set db = New ADODB.Connection
  'db.CursorLocation = adUseClient
  'db.Open gblADOProvider
  
  'Dim rsGLCOA As ADODB.Recordset
  db.Execute "UPDATE [GL Chart Of Accounts] SET [GL COA CY Beginning Amt]=0,[GL COA CY Period 1 Amt]=0," & _
  "[GL COA CY Period 2 Amt]=0,[GL COA CY Period 3 Amt]=0,[GL COA CY Period 3 Amt]=0," & _
  "[GL COA CY Period 4 Amt]=0,[GL COA CY Period 5 Amt]=0,[GL COA CY Period 6 Amt]=0," & _
  "[GL COA CY Period 7 Amt]=0,[GL COA CY Period 8 Amt]=0,[GL COA CY Period 9 Amt]=0," & _
  "[GL COA CY Period 10 Amt]=0,[GL COA CY Period 11 Amt]=0,[GL COA CY Period 12 Amt]=0," & _
  "[GL COA CY Period 13 Amt]=0,[GL COA CY Beginning Amt]=0", , adCmdText


  'If rsGLCOA.RecordCount > 0 Then
  '  rsGLCOA.MoveFirst
  '  Do While rsGLCOA.EOF = False
      ' Reset Balances
       ' current year
  '      rsGLCOA("GL COA CY Beginning Amt") = 0
  '      rsGLCOA("GL COA CY Period 1 Amt") = 0
  '      rsGLCOA("GL COA CY Period 2 Amt") = 0
  '      rsGLCOA("GL COA CY Period 3 Amt") = 0
  '      rsGLCOA("GL COA CY Period 4 Amt") = 0
  '      rsGLCOA("GL COA CY Period 5 Amt") = 0
  '      rsGLCOA("GL COA CY Period 6 Amt") = 0
  '      rsGLCOA("GL COA CY Period 7 Amt") = 0
  '      rsGLCOA("GL COA CY Period 8 Amt") = 0
  '      rsGLCOA("GL COA CY Period 9 Amt") = 0
  '      rsGLCOA("GL COA CY Period 10 Amt") = 0
  '      rsGLCOA("GL COA CY Period 11 Amt") = 0
  '      rsGLCOA("GL COA CY Period 12 Amt") = 0
  '      rsGLCOA("GL COA CY Period 13 Amt") = 0
  '      rsGLCOA("GL COA CY Period 14 Amt") = 0
  '      rsGLCOA("GL COA Account Balance") = 0

  '    rsGLCOA.Update
      ' end of Reset Balances
  '    rsGLCOA.MoveNext
  '    If Err = 3021 Then Exit Do
  '  Loop
  'End If

  'rsGLCOA.Close
  'Set rsGLCOA = Nothing
  'db.Close
  'Set db = Nothing

  Call ReloadGLBalances(db)
ShowStatus False

End Sub

Sub RollCOA(db As ADODB.Connection)
ShowStatus True
  Dim Amount@
  Dim X%
  
  'On Error GoTo RollCOA_Error

  'Put all COA period information from current year into previous year
  
  'Dim db As ADODB.Connection
  'Set db = New ADODB.Connection
  'db.CursorLocation = adUseClient
  'db.Open gblADOProvider
  
  Dim rsGLCOA As ADODB.Recordset
  Set rsGLCOA = New ADODB.Recordset
  rsGLCOA.Open "SELECT * FROM [GL Chart Of Accounts]", db, adOpenKeyset, adLockOptimistic, adCmdText
  If rsGLCOA.RecordCount = 0 Then Exit Sub
  
  db.BeginTrans

  'rsGLCOA.Index = "PrimaryKey"
  rsGLCOA.MoveFirst
  Do While Not rsGLCOA.EOF
    'Do Beginning Balance Amount
    Amount@ = IIf(IsNull(rsGLCOA("GL COA CY Beginning Amt")), 0, rsGLCOA("GL COA CY Beginning Amt"))
      rsGLCOA("GL COA PY Beginning Amt") = Amount@
    rsGLCOA.Update

    'Now do all periods
    For X% = 1 To 13
      Amount@ = IIf(IsNull(rsGLCOA("GL COA CY Period " & Trim(Str(X%)) & " Amt")), 0, rsGLCOA("GL COA CY Period " & Trim(Str(X%)) & " Amt"))
        rsGLCOA("GL COA PY Period " & Trim(Str(X%)) & " Amt") = Amount@
      rsGLCOA.Update
    Next X%
    rsGLCOA.MoveNext
    'MsgBox rsGLCOA.AbsolutePosition
  Loop
    
  db.CommitTrans

  rsGLCOA.Close
  Set rsGLCOA = Nothing
  'db.Close
  'Set db = Nothing
ShowStatus False

  Exit Sub

RollCOA_Error:
  Call ErrorLog("Year End Module", "RollCOA", Now, Err.Number, Err.Description, True, db)
  Resume Next

End Sub

Sub RollCOABalances(db As ADODB.Connection)
ShowStatus True
  'Create transactions to copy all Balance Sheet balances
  ' from Last Year Balance to Beg Balance this year
  
  Dim rs As ADODB.Recordset
  Dim sql$
  Dim Balance@
  Dim TranDate As Variant
  Dim NewNumber&
  Dim X%
  Dim Y%
  Dim BalanceType$
  Dim RetType$
  
  'Dim db As ADODB.Connection
  'Set db = New ADODB.Connection
  'db.CursorLocation = adUseClient
  'db.Open gblADOProvider

  'On Error GoTo RollCOABalances_Error

  TranDate = LookRecord("[SYS COM Fiscal End Date]", "SYS Company", db)

  sql$ = "SELECT * FROM [GL Chart of Accounts] "
  sql$ = sql$ & "where [GL COA Asset Type] <> 'Sales' AND "
  sql$ = sql$ & "[GL COA Asset Type] <> 'Cost of Goods Sold' AND "
  sql$ = sql$ & "[GL COA Asset Type] <> 'Expense'"

  rs.Open sql$, db, adOpenStatic, adLockOptimistic, adCmdText

  Dim rsGLCOA As ADODB.Recordset
  Set rsGLCOA = New ADODB.Recordset
  rsGLCOA.Open "SELECT [GL COA Balance Type] FROM [GL Chart Of Accounts]", db, adOpenKeyset, adLockOptimistic, adCmdText

  'rsGLCOA.Index = "PrimaryKey"
  'rsGLCOA.Seek gREA$
  RetType$ = rsGLCOA("GL COA Balance Type")
  rsGLCOA.Close
  Set rsGLCOA = Nothing

  ' clear any GL Work records
  'Dim cmdtemp As ADODB.Recordset
  db.Execute "DELETE DISTINCTROW * FROM [GL Work Detail]", , adCmdText
  'cmdtemp.Close
  'Set cmdtemp = Nothing

  'Date this last day of fiscal year
  TranDate = gLastDayOfYear

  'Dim rsGLTrans As ADODB.Recordset
  'rsGLTrans.Open "GL Transaction", db, adOpenStatic, adLockOptimistic, adCmdTable

  'Dim rsGLWorkDetail As ADODB.Recordset
  'rsGLWorkDetail.Open "GL Work Detail", db, adOpenStatic, adLockOptimistic, adCmdTable
  
  'Write a transaction header
SQLstatement = "INSERT INTO [GL Transaction]"
SQLstatement = SQLstatement & " ([GL TRANS Document #],[GL TRANS Type],[GL TRANS Date],"
SQLstatement = SQLstatement & " [GL TRANS Reference],[GL TRANS Amount],[GL TRANS Posted YN],"
SQLstatement = SQLstatement & " [GL TRANS Description],[GL TRANS Source],[GL TRANS System Generated])"

SQLstatement = SQLstatement & " VALUES ('APT " & AppLoginName & "','APT',#" & FormatDate(CDate(TranDate)) & "#,"
SQLstatement = SQLstatement & "'Year End'," & Balance@ & ",1,"
SQLstatement = SQLstatement & "'Year End','Year End',True)"
db.Execute SQLstatement, , adCmdText
  
  Dim rsGLTrans As ADODB.Recordset
  Set rsGLTrans = New ADODB.Recordset
  rsGLTrans.Open "SELECT [GL TRANS Number] FROM [GL Transaction] WHERE [GL TRANS Document #]='APT " & AppLoginName & "'", db, adOpenKeyset, adLockReadOnly, adCmdText
      NewNumber& = rsGLTrans("GL TRANS Number")
      rsGLTrans("GL Trans Document #") = "APT " & Trim(CStr(NewNumber&))
      rsGLTrans.Update
  rsGLTrans.Close
  Set rsGLTrans = Nothing
  'rsGLTrans.AddNew
  '  NewNumber& = rsGLTrans("GL TRANS Number")
  '  rsGLTrans("GL Trans Type") = "APT"
  '  rsGLTrans("GL Trans Document #") = "APT " & Trim(CStr(NewNumber&))
  '  rsGLTrans("GL Trans Date") = Format(TranDate, "Short Date")
  '  rsGLTrans("GL Trans Description") = "Year End"
  '  rsGLTrans("GL Trans Amount") = Balance@
  '  rsGLTrans("GL Trans Reference") = "Year End"
  '  rsGLTrans("GL Trans Source") = "Year End"
  '  rsGLTrans("GL Trans Posted YN") = 1
  '  rsGLTrans("GL TRANS System Generated") = True
  'rsGLTrans.Update

  'On Error Resume Next
  rs.MoveFirst
  Y% = 0
  Do While Not rs.EOF
    'Is account a debit or credit account?
    BalanceType$ = rs("GL COA Balance Type")
    Balance@ = 0
    Balance@ = IIf(IsNull(rs("GL COA CY Beginning Amt")), 0, rs("GL COA CY Beginning Amt"))
    For X% = 1 To 13
      Balance@ = Balance@ + rs("GL COA CY Period " & Trim(Str(X%)) & " Amt")
    Next X%
    If Balance@ = 0 Then
    Else
      If UCase$(BalanceType$) = "CREDIT" Then
        Balance@ = Balance@ * -1
      SQLstatement = "INSERT INTO [GL Work Detail]"
      SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
      SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & rs("GL COA Account No") & "',0," & Balance@ & ")"
      db.Execute SQLstatement
        'rsGLWorkDetail.AddNew
        '  rsGLWorkDetail("GW TRANSD Number") = NewNumber&
        '  rsGLWorkDetail("GW TRANSD Account") = rs("GL COA Account No")
        '  rsGLWorkDetail("GW TRANSD Debit Amount") = 0
        '  rsGLWorkDetail("GW TRANSD Credit Amount") = Balance@
        '  rsGLWorkDetail("GW TRANSD Project") = ""
        'rsGLWorkDetail.Update
      Else
      SQLstatement = "INSERT INTO [GL Work Detail]"
      SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
      SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & rs("GL COA Account No") & "'," & Balance@ & ",0)"
      db.Execute SQLstatement
        'rsGLWorkDetail.AddNew
        '  rsGLWorkDetail("GW TRANSD Number") = NewNumber&
        '  rsGLWorkDetail("GW TRANSD Account") = rs("GL COA Account No")
        '  rsGLWorkDetail("GW TRANSD Debit Amount") = Balance@
        '  rsGLWorkDetail("GW TRANSD Credit Amount") = 0
        '  rsGLWorkDetail("GW TRANSD Project") = ""
        'rsGLWorkDetail.Update
      End If
    End If
    rs.MoveNext
    Y% = Y% + 1
  Loop

  Dim Success%
  Success% = PostGLWorkDetail(TranDate, NewNumber&, db)
  
  rs.Close
  Set rs = Nothing
  'rsGLWorkDetail.Close
  'Set rsGLWorkDetail = Nothing
  'rsGLTrans.Close
  'Set rsGLTrans = Nothing
  'rsGLCOA.Close
  'Set rsGLCOA = Nothing
  'db.Close
  'Set db = Nothing
ShowStatus False

  Exit Sub

RollCOABalances_Error:
  Call ErrorLog("Year End Module", "RollCOABalances", Now, Err.Number, Err.Description, True, db)
  Resume Next

End Sub

Sub VendorRollOver(VendorID$, RollDate As Variant, db As ADODB.Connection)
ShowStatus True
  'On Error Resume Next

  'Dim db As ADODB.Connection
  'Set db = New ADODB.Connection
  'db.CursorLocation = adUseClient
  'db.Open gblADOProvider
  Dim sql$
  Dim rs As ADODB.Recordset

  'Do a year end roll over for this vendor
  'Need to Roll Over:
  '   Purchase Amount
  '   Payment Amount
  '   1099 Payment Amount
  '   Number of Purchases
  '   Number of Payments

  'Find this vendor record
  Dim rsVendor As ADODB.Recordset
  Set rsVendor = New ADODB.Recordset
  rsVendor.Open "SELECT [AP VEN ID],[AP VEN Purchase YTD],[AP VEN Purchase Last Year]," & _
  "[AP VEN Purchase Number YTD],[AP VEN Purchase Number Last Year]," & _
  "[AP VEN Payments YTD],[AP VEN Payments Last Year],[AP VEN Payment Number YTD]," & _
  "[AP VEN Payment Number Last Year] FROM [AP Vendor] WHERE [AP VEN ID]='" & VendorID$ & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
  'rsVendor.Index = "PrimaryKey"
  'rsVendor.Seek VendorID$
  If rsVendor.RecordCount = 0 Then
    Exit Sub
  End If

  'Do Purchase first
  Dim PurchaseYTD#
  Dim TempPurchaseYTD#
  Dim PurchaseNumberYTD%
  Dim TempPurchaseNumberYTD%

  Dim DayOne As Variant
  DayOne = LookRecord("[SYS COM Fiscal Start Date]", "[SYS Company]", db)

  Dim Purchases#
  Dim Refunds#
  Purchases# = IIf(IsNull(SumRecord("[AP PO Total Amount]", "[AP Purchase]", db, "[AP PO Vendor ID] = '" & VendorID$ & "' AND [AP PO Posted YN] = TRUE AND [AP PO Date] >= #" & DayOne & "# AND [AP PO Document Type] in ('Receiving','Voucher','Beginning Balance')")), 0, SumRecord("[AP PO Total Amount]", "AP Purchase", db, "[AP PO Vendor ID] = '" & gVendorID$ & "' AND [AP PO Posted YN] = TRUE AND [AP PO Date] >= #" & DayOne & "# AND [AP PO Document Type] in ('Receiving','Voucher','Beginning Balance')"))
  Refunds# = IIf(IsNull(SumRecord("[AP PO Total Amount]", "[AP Purchase]", db, "[AP PO Vendor ID] = '" & VendorID$ & "' AND [AP PO Posted YN] = TRUE AND [AP PO Date] >= #" & DayOne & "# AND [AP PO Document Type] in ('Credit Memo')")), 0, SumRecord("[AP PO Total Amount]", "AP Purchase", db, "[AP PO Vendor ID] = '" & gVendorID$ & "' AND [AP PO Posted YN] = TRUE AND [AP PO Date] >= #" & DayOne & "# AND [AP PO Document Type] in ('Credit Memo')"))
  PurchaseYTD# = Purchases# - Refunds#
  TempPurchaseYTD# = PurchaseYTD#

  PurchaseNumberYTD% = CountRecord("[AP PO Total Amount]", "[AP Purchase]", db, "[AP PO Vendor ID] = '" & VendorID$ & "' AND [AP PO Posted YN] = TRUE AND [AP PO Date] >= #" & DayOne & "# AND [AP PO Document Type] in ('Receiving','Voucher','Beginning Balance')")
  TempPurchaseNumberYTD% = PurchaseNumberYTD%

  sql$ = "SELECT [AP PO Total Amount] FROM [AP Purchase] WHERE [AP PO Vendor ID] = '" & VendorID$ & "' AND [AP PO Date] > #" & RollDate & "#"
  Set rs = New ADODB.Recordset
  rs.Open sql$, db, adOpenStatic, adLockOptimistic, adCmdText
  
  'On Error Resume Next
  'Err = 0
  If rs.RecordCount > 0 Then
  rs.MoveFirst
    Do While Not rs.EOF
      TempPurchaseYTD# = TempPurchaseYTD# - rs("AP PO Total Amount")
      TempPurchaseNumberYTD% = TempPurchaseNumberYTD% - 1
      rs.MoveNext
    Loop
  End If
  rs.Close
  Set rs = Nothing
  
  'Write out new data to vendor file
    rsVendor("AP VEN Purchase YTD") = PurchaseYTD# - TempPurchaseYTD#
    rsVendor("AP VEN Purchase Last Year") = TempPurchaseYTD#
    rsVendor("AP VEN Purchase Number YTD") = PurchaseNumberYTD% - TempPurchaseNumberYTD%
    rsVendor("AP VEN Purchase Number Last Year") = TempPurchaseNumberYTD%
  rsVendor.Update
  '---------------------------------------------------------------------
  'Now do Payments
  
  Dim PaymentsYTD#
  Dim TempPaymentsYTD#
  Dim NumberPaymentsYTD%
  Dim TempNumberPaymentsYTD%

  PaymentsYTD# = SumRecord("[AP PAY Amount]", "[AP Payment Header]", db, "[AP PAY Vendor No] = '" & VendorID$ & "' AND [AP PAY Posted YN] = TRUE AND [AP PAY Transaction Date] >= #" & DayOne & "# AND [AP PAY Void] = FALSE AND [AP PAY Type] <> 'Credit Memo'")
  TempPaymentsYTD# = PaymentsYTD#

  NumberPaymentsYTD% = CountRecord("[AP PAY Amount]", "[AP Payment Header]", db, "[AP PAY Vendor No] = '" & VendorID$ & "' AND [AP PAY Posted YN] = TRUE AND [AP PAY Transaction Date] >= #" & DayOne & "# AND [AP PAY Void] = FALSE AND [AP PAY Type] <> 'Credit Memo'")
  TempNumberPaymentsYTD% = NumberPaymentsYTD%
  

  sql$ = "SELECT [AP PAY Amount] FROM [AP Payment Header] WHERE [AP PAY Vendor No] = '" & VendorID$ & "' AND [AP PAY Transaction Date] > #" & RollDate & "#"
    
  Dim rs2 As ADODB.Recordset
  Set rs2 = New ADODB.Recordset
  rs2.Open sql$, db, adOpenStatic, adLockOptimistic, adCmdText

  'On Error Resume Next
  'Err = 0
  If rs2.RecordCount > 0 Then
    rs2.MoveFirst
    Do While Not rs2.EOF
      TempPaymentsYTD# = TempPaymentsYTD# - rs2("AP PAY Amount")
      TempNumberPaymentsYTD% = TempNumberPaymentsYTD% - 1
      rs2.MoveNext
    Loop
  End If
  rs2.Close
  Set rs2 = Nothing
  
  'Write out new data to vendor file
    rsVendor("AP VEN Payments YTD") = PaymentsYTD# - TempPaymentsYTD#
    rsVendor("AP VEN Payments Last Year") = TempPaymentsYTD#
    rsVendor("AP VEN Payment Number YTD") = NumberPaymentsYTD% - TempNumberPaymentsYTD%
    rsVendor("AP VEN Payment Number Last Year") = TempNumberPaymentsYTD%
  rsVendor.Update


'rs.Close
'Set rs = Nothing
'rs2.Close
'Set rs2 = Nothing
rsVendor.Close
Set rsVendor = Nothing
'db.Close
'Set db = Nothing
ShowStatus False

End Sub

Sub WriteRetainedEarnings(db As ADODB.Connection)
ShowStatus True
  'On Error Resume Next

  'Do a transaction for all P&L Accounts to zero their balance and out
  '  difference into retained earnings account
  'Dim db As ADODB.Connection
  'Set db = New ADODB.Connection
  'db.CursorLocation = adUseClient
  'db.Open gblADOProvider
  
  Dim rs As ADODB.Recordset
  Dim sql$
  Dim Balance@
  Dim RetBalance@
  Dim BalanceType$
  Dim NewNumber&
  ReDim AcctNum$(0 To 100)
  ReDim AcctName$(0 To 100)
  ReDim BalType$(0 To 100)
  ReDim PrevBal@(0 To 100)
  Dim RetName$
  Dim X%
  Dim RetType$
  Dim Y%
  Dim SQLstatement As String
  
  Dim rsGLCOA As ADODB.Recordset
  Set rsGLCOA = New ADODB.Recordset
  
  SQLstatement = ""
  For X% = 1 To 14
    SQLstatement = SQLstatement + "[GL COA CY Period " & Trim(Str(X%)) & " Amt],"
  Next X%
  
  rsGLCOA.Open "SELECT " & SQLstatement & " [GL COA Balance Type],[GL COA Account Name] " & _
  "FROM [GL Chart Of Accounts] WHERE [GL COA Account No]='" & gREA$ & "'", db, adOpenKeyset, adLockOptimistic, adCmdText

  'rsGLCOA.Index = "PrimaryKey"
  'rsGLCOA.Seek gREA$
  RetName$ = rsGLCOA("GL COA Account Name")
  RetType$ = rsGLCOA("GL COA Balance Type")

  'What is balance of this account
  Balance@ = 0
  For X% = 1 To 14
    Balance@ = Balance@ + rsGLCOA("GL COA CY Period " & Trim(Str(X%)) & " Amt")
  Next X%
  RetBalance@ = Balance@

  rsGLCOA.Close
  Set rsGLCOA = Nothing

  sql$ = "SELECT * FROM [GL Chart of Accounts] "
  sql$ = sql$ & "where [GL COA Asset Type] = 'Sales' OR "
  sql$ = sql$ & "[GL COA Asset Type] = 'Cost of Goods Sold' OR "
  sql$ = sql$ & "[GL COA Asset Type] = 'Expense'"
  
  Set rs = New ADODB.Recordset
  rs.Open sql$, db, adOpenKeyset, adLockOptimistic, adCmdText
  If rs.RecordCount = 0 Then Exit Sub
  ' clear any GL Work records
  'Dim cmdtemp As ADODB.Recordset
  db.Execute "DELETE DISTINCTROW * FROM [GL Work Detail]", , adCmdText
  'cmdtemp.Close
  'Set cmdtemp = Nothing

  
  'Date this last day of fiscal year
  Dim TranDate As Variant
  TranDate = gLastDayOfYear
  
  'Dim rsGLTrans As ADODB.Recordset
  'rsGLTrans.Open "GL Transaction", db, adOpenStatic, adLockOptimistic, adCmdTable
  
  'Write a transaction header
SQLstatement = "INSERT INTO [GL Transaction]"
SQLstatement = SQLstatement & " ([GL TRANS Document #],[GL TRANS Type],[GL TRANS Date],"
SQLstatement = SQLstatement & " [GL TRANS Reference],[GL TRANS Amount],[GL TRANS Posted YN],"
SQLstatement = SQLstatement & " [GL TRANS Description],[GL TRANS Source],[GL TRANS System Generated])"

SQLstatement = SQLstatement & " VALUES ('BEGBAL " & AppLoginName & "','BEG BAL',#" & FormatDate(CDate(TranDate)) & "#,"
SQLstatement = SQLstatement & "'Year End'," & Balance@ & ",1,"
SQLstatement = SQLstatement & "'Year End','Year End',True)"
db.Execute SQLstatement, , adCmdText
  
  Dim rsGLTrans As ADODB.Recordset
  Set rsGLTrans = New ADODB.Recordset
  rsGLTrans.Open "SELECT [GL TRANS Number],[GL Trans Document #] FROM [GL Transaction] WHERE [GL TRANS Document #]='BEGBAL " & AppLoginName & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
      NewNumber& = rsGLTrans("GL TRANS Number")
      rsGLTrans("GL Trans Document #") = "BEG BAL " & Trim(CStr(NewNumber&))
      rsGLTrans.Update
  rsGLTrans.Close
  Set rsGLTrans = Nothing
  
  'rsGLTrans.AddNew
    'NewNumber& = rsGLTrans("GL TRANS Number")
    'rsGLTrans("GL Trans Document #") = "BEGBAL " & Trim(CStr(NewNumber&))
    'rsGLTrans("GL Trans Type") = "BEGBAL"
    'rsGLTrans("GL Trans Date") = Format(TranDate, "Short Date")
    'rsGLTrans("GL Trans Reference") = "Year End"
    'rsGLTrans("GL Trans Amount") = Balance@
    'rsGLTrans("GL Trans Posted YN") = 1
    'rsGLTrans("GL Trans Description") = "Year End"
    'rsGLTrans("GL Trans Source") = "Year End"
    'rsGLTrans("GL TRANS System Generated") = True
  'rsGLTrans.Update

  'Dim rsGLWorkDetail As ADODB.Recordset
  'rsGLWorkDetail.Open "GL Work Detail", db, adOpenStatic, adLockOptimistic, adCmdTable

  'On Error Resume Next
  rs.MoveFirst
  Y% = 0
  RetBalance@ = 0
  Do While Not rs.EOF
    'Is account a debit or credit account?
    BalanceType$ = rs("GL COA Balance Type")
    Balance@ = 0
    Balance@ = IIf(IsNull(rs("GL COA CY Beginning Amt")), 0, rs("GL COA CY Beginning Amt"))
    For X% = 1 To 13
      Balance@ = Balance@ + rs("GL COA CY Period " & Trim(Str(X%)) & " Amt")
    Next X%
    If Balance@ <> 0 Then
      AcctNum$(Y%) = rs("GL COA Account No")
      AcctName$(Y%) = rs("GL COA Account Name")
      RetBalance@ = RetBalance@ + Balance@
      PrevBal@(Y%) = Balance@
      'Balance@ = Abs(Balance@)
      If UCase$(BalanceType$) = "CREDIT" Then
        Balance@ = Balance@ * -1
        BalType$(Y%) = "Credit"
        
      SQLstatement = "INSERT INTO [GL Work Detail]"
      SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
      SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & rs("GL COA Account No") & "'," & Balance@ & ",0)"
      db.Execute SQLstatement
      '  rsGLWorkDetail.AddNew
      '    rsGLWorkDetail("GW TRANSD Number") = NewNumber&
      '    rsGLWorkDetail("GW TRANSD Account") = rs("GL COA Account No")
      '    rsGLWorkDetail("GW TRANSD Debit Amount") = Balance@
      '    rsGLWorkDetail("GW TRANSD Credit Amount") = 0
      '    rsGLWorkDetail("GW TRANSD Project") = ""
      '  rsGLWorkDetail.Update
      SQLstatement = "INSERT INTO [GL Work Detail]"
      SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
      SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & gREA$ & "',0," & Balance@ & ")"
      db.Execute SQLstatement
      '  rsGLWorkDetail.AddNew
      '    rsGLWorkDetail("GW TRANSD Number") = NewNumber&
      '    rsGLWorkDetail("GW TRANSD Account") = gREA$
      '    rsGLWorkDetail("GW TRANSD Debit Amount") = 0
      '    rsGLWorkDetail("GW TRANSD Credit Amount") = Balance@
      '    rsGLWorkDetail("GW TRANSD Project") = ""
      '  rsGLWorkDetail.Update
      Else
        BalType$(Y%) = "Debit"
      SQLstatement = "INSERT INTO [GL Work Detail]"
      SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
      SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & rs("GL COA Account No") & "',0," & Balance@ & ")"
      db.Execute SQLstatement
      '  rsGLWorkDetail.AddNew
      '    rsGLWorkDetail("GW TRANSD Number") = NewNumber&
      '    rsGLWorkDetail("GW TRANSD Account") = rs("GL COA Account No")
      '    rsGLWorkDetail("GW TRANSD Debit Amount") = 0
      '    rsGLWorkDetail("GW TRANSD Credit Amount") = Balance@
      '    rsGLWorkDetail("GW TRANSD Project") = ""
      '  rsGLWorkDetail.Update
      SQLstatement = "INSERT INTO [GL Work Detail]"
      SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
      SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & rs("GL COA Account No") & "'," & Balance@ & ",0)"
      db.Execute SQLstatement
      '  rsGLWorkDetail.AddNew
      '    rsGLWorkDetail("GW TRANSD Number") = NewNumber&
      '    rsGLWorkDetail("GW TRANSD Account") = gREA$
      '    rsGLWorkDetail("GW TRANSD Debit Amount") = Balance@
      '    rsGLWorkDetail("GW TRANSD Credit Amount") = 0
      '    rsGLWorkDetail("GW TRANSD Project") = ""
      '  rsGLWorkDetail.Update
      End If
      Y% = Y% + 1
    End If
    rs.MoveNext
  Loop
  'rsGLWorkDetail.Close
  

  Dim Success%
  Success% = PostGLWorkDetail(TranDate, NewNumber&, db)

  'On Error Resume Next

rs.Close
Set rs = Nothing
ShowStatus False
'rsGLWorkDetail.Close
'Set GLWorkDetail = Nothing
'rsGLTrans.Close
'Set rsGLTrans = Nothing
'db.Close
'Set db = Nothing
End Sub

