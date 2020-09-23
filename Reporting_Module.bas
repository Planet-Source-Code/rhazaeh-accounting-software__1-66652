Attribute VB_Name = "GL_Report_Module"

Function BuildWorkBalanceSheet(LEVEL%, ThroughPeriod%)

  'On Error GoTo BuildWorkBalanceSheet_Error
  
  Dim db As ADODB.Connection
  Set db = New ADODB.Connection
  db.CursorLocation = adUseClient
  db.Open gblADOProvider
  
  Dim x%
  
  Dim rsGL As ADODB.Recordset
  rsGL.Open "GL Chart Of Accounts", db, adOpenStatic, adLockOptimistic, adCmdTable
  
  Dim cmdtemp As ADODB.Recordset
  cmdtemp.Open "DELETE * FROM [Work - Balance Sheet]", db, , , adCmdText
  cmdtemp.Close
  Set cmdtemp = Nothing
    
  Dim rsWork As ADODB.Recordset
  rsWork.Open "Work - Balance Sheet", db, adOpenStatic, adLockOptimistic, adCmdTable

  Dim BalanceType$
  ' AssetType$ added to check for Acc Depr type to maintain value as credit xxx 8/22/97
  Dim AssetType$

  rsGL.Index = "PrimaryKey"

  rsGL.MoveFirst
  Dim TempCYBalance@
  Dim TempPYBalance@
  Dim TempBUDBalance@
  Do While Not rsGL.EOF
    TempCYBalance@ = rsGL("GL COA CY Beginning Amt")
    TempPYBalance@ = rsGL("GL COA PY Beginning Amt")
    TempBUDBalance@ = rsGL("GL COA Bud Beginning Amt")
    For x% = 1 To ThroughPeriod%
      TempCYBalance@ = TempCYBalance@ + rsGL("GL COA CY Period " & Trim(CStr(x%)) & " Amt")
      TempPYBalance@ = TempPYBalance@ + rsGL("GL COA PY Period " & Trim(CStr(x%)) & " Amt")
      TempBUDBalance@ = TempBUDBalance@ + rsGL("GL COA Bud Period " & Trim(CStr(x%)) & " Amt")
    Next x%

    'Add it to the work table
    rsWork.AddNew
      If rsGL("GL COA Reporting Level") <= LEVEL% Then
        rsWork("Visible") = True
      Else
        rsWork("Visible") = False
      End If
      BalanceType$ = rsGL("GL COA Balance Type")
      AssetType$ = rsGL("GL COA Asset Type")
      rsWork("CY Balance") = IIf(BalanceType$ = "Debit" Or AssetType$ = "Accum Depreciation", TempCYBalance@, TempCYBalance@ * -1)
      rsWork("PY Balance") = IIf(BalanceType$ = "Debit" Or AssetType$ = "Accum Depreciation", TempPYBalance@, TempPYBalance@ * -1)
      rsWork("BUD Balance") = IIf(BalanceType$ = "Debit", TempBUDBalance@, TempBUDBalance@ * -1)
      rsWork("GL COA Account No") = rsGL("GL COA Account No")
      rsWork("GL COA Account Name") = rsGL("GL COA Account Name")
      rsWork("GL COA Reporting Level") = rsGL("GL COA Reporting Level")
      rsWork("GL COA Balance Type") = rsGL("GL COA Balance Type")
      rsWork("GL COA Asset Type") = rsGL("GL COA Asset Type")
    rsWork.Update
    
    rsGL.MoveNext
  Loop

  'Add it to the work table
  rsWork.AddNew
    rsWork("Visible") = True
    BalanceType$ = "Credit"
    rsWork("CY Balance") = ComputeNetProfit("CY", ThroughPeriod%, db)
    rsWork("PY Balance") = ComputeNetProfit("PY", ThroughPeriod%, db)
    rsWork("BUD Balance") = ComputeNetProfit("Budget", ThroughPeriod%, db)
    rsWork("GL COA Account No") = "399999"
    rsWork("GL COA Account Name") = "Current Earnings"
    rsWork("GL COA Reporting Level") = 1
    rsWork("GL COA Balance Type") = "Credit"
    rsWork("GL COA Asset Type") = "Equity"
  rsWork.Update

  'Calculate total assets
  Dim TotalCYAssets@
  Dim TotalPYAssets@
  Dim TotalBUDAssets@

  TotalCYAssets@ = SumRecord("[CY Balance]", "[Work - Balance Sheet]", db, "[GL COA Asset Type] in ('Accounts Receivable','Accum Depreciation','Cash','Fixed Assets','Inventory','Other Assets','Other Current Assets')")
  TotalPYAssets@ = SumRecord("[PY Balance]", "[Work - Balance Sheet]", db, "[GL COA Asset Type] in ('Accounts Receivable','Accum Depreciation','Cash','Fixed Assets','Inventory','Other Assets','Other Current Assets')")
  TotalBUDAssets@ = SumRecord("[BUD Balance]", "[Work - Balance Sheet]", db, "[GL COA Asset Type] in ('Accounts Receivable','Accum Depreciation','Cash','Fixed Assets','Inventory','Other Assets','Other Current Assets')")

  'Calculate Total Liabilities
  Dim TotalCYLiabilities@
  Dim TotalPYLiabilities@
  Dim TotalBUDLiabilities@
  
  TotalCYLiabilities@ = SumRecord("[CY Balance]", "[Work - Balance Sheet]", db, "[GL COA Asset Type] in ('Accounts Payable','Credit Cards Payable','Long Term Liabilities','Other Current Liabilities','Taxes Payable')")
  TotalPYLiabilities@ = SumRecord("[PY Balance]", "[Work - Balance Sheet]", db, "[GL COA Asset Type] in ('Accounts Payable','Credit Cards Payable','Long Term Liabilities','Other Current Liabilities','Taxes Payable')")
  TotalBUDLiabilities@ = SumRecord("[BUD Balance]", "[Work - Balance Sheet]", db, "[GL COA Asset Type] in ('Accounts Payable','Credit Cards Payable','Long Term Liabilities','Other Current Liabilities','Taxes Payable')")
  
  'Calculate Total Equity
  Dim TotalCYEquity@
  Dim TotalPYEquity@
  Dim TotalBUDEquity@
  
  TotalCYEquity@ = SumRecord("[CY Balance]", "[Work - Balance Sheet]", db, "[GL COA Asset Type] in ('Equity')")
  TotalPYEquity@ = SumRecord("[PY Balance]", "[Work - Balance Sheet]", db, "[GL COA Asset Type] in ('Equity')")
  TotalBUDEquity@ = SumRecord("[BUD Balance]", "[Work - Balance Sheet]", db, "[GL COA Asset Type] in ('Equity')")

  'Roll balances
  Dim HoldCYBalance@
  Dim HoldPYBalance@
  Dim HoldBUDBalance@

  rsWork.MoveFirst
  Do While Not rsWork.EOF
    HoldCYBalance@ = rsWork("CY Balance")
    HoldPYBalance@ = rsWork("PY Balance")
    HoldBUDBalance@ = rsWork("BUD Balance")
  
    Dim rs As ADODB.Recordset
    rs.Open "SELECT * FROM [Work - Balance Sheet] ORDER BY [GL COA Account No]", db, adOpenStatic, adLockOptimistic, adCmdText
    rs.MoveFirst
    rs.Find "[GL COA Account No] = '" & Trim(CStr(rsWork("GL COA Account No"))) & "'"
  
    If rs("GL COA Reporting Level") > LEVEL% Then
      HoldCYBalance@ = 0
      HoldPYBalance@ = 0
      HoldBUDBalance@ = 0
    Else
      'On Error Resume Next
      rs.MoveNext
      If Err = 0 Then
        Do While rs("GL COA Reporting Level") > LEVEL% And Not rs.EOF
          HoldCYBalance@ = HoldCYBalance@ + rs("CY Balance")
          HoldPYBalance@ = HoldPYBalance@ + rs("PY Balance")
          HoldBUDBalance@ = HoldBUDBalance@ + rs("BUD Balance")
          rs.MoveNext
          If Err = 3021 Then Exit Do
        Loop
      End If
    End If

      rsWork("CY Balance") = HoldCYBalance@
      rsWork("CY Ratio") = IIf(TotalCYAssets@ = 0, 0, HoldCYBalance@ / TotalCYAssets@)
      rsWork("PY Balance") = HoldPYBalance@
      rsWork("PY Ratio") = IIf(TotalPYAssets@ = 0, 0, HoldPYBalance@ / TotalPYAssets@)
      rsWork("BUD Balance") = HoldBUDBalance@
      rsWork("BUD Ratio") = IIf(TotalBUDAssets@ = 0, 0, HoldBUDBalance@ / TotalBUDAssets@)
      rsWork("Total CY Assets") = TotalCYAssets@
      rsWork("Total PY Assets") = TotalPYAssets@
      rsWork("Total BUD Assets") = TotalBUDAssets@
      rsWork("Total CY Liabilities") = TotalCYLiabilities@
      rsWork("Total PY Liabilities") = TotalPYLiabilities@
      rsWork("Total BUD Liabilities") = TotalBUDLiabilities@
      rsWork("Total CY Equity") = TotalCYEquity@
      rsWork("Total PY Equity") = TotalPYEquity@
      rsWork("Total BUD Equity") = TotalBUDEquity@
    rsWork.Update
      
    rsWork.MoveNext
  Loop

  rs.Close
  Set rs = Nothing
  rsGL.Close
  Set rsGL = Nothing
  rsWork.Close
  Set rsWork = Nothing
  db.Close
  Set db = Nothing
  
  Exit Function
BuildWorkBalanceSheet_Error:
  Call ErrorLog("Reporting Module", "BuildWorkBalanceSheet", Now, Err.Number, Err.Description, True, db)
  
  rs.Close
  Set rs = Nothing
  rsGL.Close
  Set rsGL = Nothing
  rsWork.Close
  Set rsWork = Nothing
  db.Close
  Set db = Nothing

End Function

Function BuildWorkIncomeStatement(LEVEL%, StartPeriod%, EndPeriod%)

  'On Error GoTo BuildWorkIncomeStatement_Error

  Dim x%
  Dim db As ADODB.Connection
  Set db = New ADODB.Connection
  db.CursorLocation = adUseClient
  db.Open gblADOProvider
  
  Dim cmdtemp As ADODB.Recordset
  cmdtemp.Open "DELETE * FROM [Work - Income Statement]", db, , , adCmdText
  cmdtemp.Close
  Set cmdtemp = Nothing

  Dim rsGL As ADODB.Recordset
  rsGL.Open "GL Chart Of Accounts", db, adOpenStatic, adLockOptimistic, adCmdTable
  
  Dim rsWork As ADODB.Recordset
  rsWork.Open "Work - Income Statement", db, adOpenStatic, adLockOptimistic, adCmdTable

  Dim BalanceType$

  rsGL.Index = "PrimaryKey"

  rsGL.MoveFirst
  Dim TempCYPeriodBalance@
  Dim TempPYPeriodBalance@
  Dim TempBUDPeriodBalance@
  Dim TempCYYTDBalance@
  Dim TempPYYTDBalance@
  Dim TempBUDYTDBalance@
  Do While Not rsGL.EOF
    TempCYPeriodBalance@ = 0
    TempPYPeriodBalance@ = 0
    TempBUDPeriodBalance@ = 0
    TempCYYTDBalance@ = rsGL("GL COA CY Beginning Amt")
    TempPYYTDBalance@ = rsGL("GL COA PY Beginning Amt")
    TempBUDYTDBalance@ = rsGL("GL COA Bud Beginning Amt")
    For x% = 1 To EndPeriod%
      TempCYYTDBalance@ = TempCYYTDBalance@ + rsGL("GL COA CY Period " & Trim(CStr(x%)) & " Amt")
      TempPYYTDBalance@ = TempPYYTDBalance@ + rsGL("GL COA PY Period " & Trim(CStr(x%)) & " Amt")
      TempBUDYTDBalance@ = TempBUDYTDBalance@ + rsGL("GL COA Bud Period " & Trim(CStr(x%)) & " Amt")
      If x% >= StartPeriod% Then
        TempCYPeriodBalance@ = TempCYPeriodBalance@ + rsGL("GL COA CY Period " & Trim(CStr(x%)) & " Amt")
        TempPYPeriodBalance@ = TempPYPeriodBalance@ + rsGL("GL COA PY Period " & Trim(CStr(x%)) & " Amt")
        TempBUDPeriodBalance@ = TempBUDPeriodBalance@ + rsGL("GL COA Bud Period " & Trim(CStr(x%)) & " Amt")
      End If
    Next x%

    'Add it to the work table
    rsWork.AddNew
      If rsGL("GL COA Reporting Level") <= LEVEL% Then
        rsWork("Visible") = True
      Else
        rsWork("Visible") = False
      End If
      BalanceType$ = rsGL("GL COA Balance Type")
      rsWork("CY Period Balance") = IIf(BalanceType$ = "Debit", TempCYPeriodBalance@, TempCYPeriodBalance@ * -1)
      rsWork("PY Period Balance") = IIf(BalanceType$ = "Debit", TempPYPeriodBalance@, TempPYPeriodBalance@ * -1)
      rsWork("BUD Period Balance") = IIf(BalanceType$ = "Debit", TempBUDPeriodBalance@, TempBUDPeriodBalance@ * -1)
      rsWork("CY YTD Balance") = IIf(BalanceType$ = "Debit", TempCYYTDBalance@, TempCYYTDBalance@ * -1)
      rsWork("PY YTD Balance") = IIf(BalanceType$ = "Debit", TempPYYTDBalance@, TempPYYTDBalance@ * -1)
      rsWork("BUD YTD Balance") = IIf(BalanceType$ = "Debit", TempBUDYTDBalance@, TempBUDYTDBalance@ * -1)
      rsWork("GL COA Account No") = rsGL("GL COA Account No")
      rsWork("GL COA Account Name") = rsGL("GL COA Account Name")
      rsWork("GL COA Reporting Level") = rsGL("GL COA Reporting Level")
      rsWork("GL COA Balance Type") = rsGL("GL COA Balance Type")
      rsWork("GL COA Asset Type") = rsGL("GL COA Asset Type")
    rsWork.Update
    
    rsGL.MoveNext
  Loop

  'Calculate totals
  Dim TotalCYPeriodIncome@
  Dim TotalPYPeriodIncome@
  Dim TotalBUDPeriodIncome@
  Dim TotalCYPeriodCOGS@
  Dim TotalPYPeriodCOGS@
  Dim TotalBUDPeriodCOGS@
  Dim TotalCYPeriodExpenses@
  Dim TotalPYPeriodExpenses@
  Dim TotalBUDPeriodExpenses@
  
  Dim TotalCYYTDIncome@
  Dim TotalPYYTDIncome@
  Dim TotalBUDYTDIncome@
  Dim TotalCYYTDCOGS@
  Dim TotalPYYTDCOGS@
  Dim TotalBUDYTDCOGS@
  Dim TotalCYYTDExpenses@
  Dim TotalPYYTDExpenses@
  Dim TotalBUDYTDExpenses@

  TotalCYPeriodIncome@ = NZ(SumRecord("[CY Period Balance]", "[Work - Income Statement]", db, "[GL COA Asset Type] = 'Sales'"), 0)
  TotalPYPeriodIncome@ = NZ(SumRecord("[PY Period Balance]", "[Work - Income Statement]", db, "[GL COA Asset Type] = 'Sales'"), 0)
  TotalBUDPeriodIncome@ = NZ(SumRecord("[BUD Period Balance]", "[Work - Income Statement]", db, "[GL COA Asset Type] = 'Sales'"), 0)
  TotalCYPeriodCOGS@ = NZ(SumRecord("[CY Period Balance]", "[Work - Income Statement]", db, "[GL COA Asset Type] = 'Cost of Goods Sold'"), 0)
  TotalPYPeriodCOGS@ = NZ(SumRecord("[PY Period Balance]", "[Work - Income Statement]", db, "[GL COA Asset Type] = 'Cost of Goods Sold'"), 0)
  TotalBUDPeriodCOGS@ = NZ(SumRecord("[BUD Period Balance]", "[Work - Income Statement]", db, "[GL COA Asset Type] = 'Cost of Goods Sold'"), 0)
  TotalCYPeriodExpenses@ = NZ(SumRecord("[CY Period Balance]", "[Work - Income Statement]", db, "[GL COA Asset Type] = 'Expense'"), 0)
  TotalPYPeriodExpenses@ = NZ(SumRecord("[PY Period Balance]", "[Work - Income Statement]", db, "[GL COA Asset Type] = 'Expense'"), 0)
  TotalBUDPeriodExpenses@ = NZ(SumRecord("[BUD Period Balance]", "[Work - Income Statement]", db, "[GL COA Asset Type] = 'Expense'"), 0)

  TotalCYYTDIncome@ = NZ(SumRecord("[CY YTD Balance]", "[Work - Income Statement]", db, "[GL COA Asset Type] = 'Sales'"), 0)
  TotalPYYTDIncome@ = NZ(SumRecord("[PY YTD Balance]", "[Work - Income Statement]", db, "[GL COA Asset Type] = 'Sales'"), 0)
  TotalBUDYTDIncome@ = NZ(SumRecord("[BUD YTD Balance]", "[Work - Income Statement]", db, "[GL COA Asset Type] = 'Sales'"), 0)
  TotalCYYTDCOGS@ = NZ(SumRecord("[CY YTD Balance]", "[Work - Income Statement]", db, "[GL COA Asset Type] = 'Cost of Goods Sold'"), 0)
  TotalPYYTDCOGS@ = NZ(SumRecord("[PY YTD Balance]", "[Work - Income Statement]", db, "[GL COA Asset Type] = 'Cost of Goods Sold'"), 0)
  TotalBUDYTDCOGS@ = NZ(SumRecord("[BUD YTD Balance]", "[Work - Income Statement]", db, "[GL COA Asset Type] = 'Cost of Goods Sold'"), 0)
  TotalCYYTDExpenses@ = NZ(SumRecord("[CY YTD Balance]", "[Work - Income Statement]", db, "[GL COA Asset Type] = 'Expense'"), 0)
  TotalPYYTDExpenses@ = NZ(SumRecord("[PY YTD Balance]", "[Work - Income Statement]", db, "[GL COA Asset Type] = 'Expense'"), 0)
  TotalBUDYTDExpenses@ = NZ(SumRecord("[BUD YTD Balance]", "[Work - Income Statement]", db, "[GL COA Asset Type] = 'Expense'"), 0)

  'Roll balances
  Dim HoldCYPeriodBalance@
  Dim HoldPYPeriodBalance@
  Dim HoldBUDPeriodBalance@
  Dim HoldCYYTDBalance@
  Dim HoldPYYTDBalance@
  Dim HoldBUDYTDBalance@

  rsWork.MoveFirst
  Do While Not rsWork.EOF
    HoldCYPeriodBalance@ = rsWork("CY Period Balance")
    HoldPYPeriodBalance@ = rsWork("PY Period Balance")
    HoldBUDPeriodBalance@ = rsWork("BUD Period Balance")
    HoldCYYTDBalance@ = rsWork("CY YTD Balance")
    HoldPYYTDBalance@ = rsWork("PY YTD Balance")
    HoldBUDYTDBalance@ = rsWork("BUD YTD Balance")
  
    Dim rs As ADODB.Recordset
    rs.Open "SELECT * FROM [Work - Income Statement] ORDER BY [GL COA Account No]", db, adOpenStatic, adLockOptimistic, adCmdText
    rs.MoveFirst
    rs.Find "[GL COA Account No] = '" & Trim(CStr(rsWork("GL COA Account No"))) & "'"
  
    If rs("GL COA Reporting Level") > LEVEL% Then
      HoldCYPeriodBalance@ = 0
      HoldPYPeriodBalance@ = 0
      HoldBUDPeriodBalance@ = 0
      HoldCYYTDBalance@ = 0
      HoldPYYTDBalance@ = 0
      HoldBUDYTDBalance@ = 0
    Else
      'On Error Resume Next
      rs.MoveNext
      If Err = 0 Then
        Do While rs("GL COA Reporting Level") > LEVEL% And Not rs.EOF
          HoldCYPeriodBalance@ = HoldCYPeriodBalance@ + rs("CY Period Balance")
          HoldPYPeriodBalance@ = HoldPYPeriodBalance@ + rs("PY Period Balance")
          HoldBUDPeriodBalance@ = HoldBUDPeriodBalance@ + rs("BUD Period Balance")
          HoldCYYTDBalance@ = HoldCYYTDBalance@ + rs("CY YTD Balance")
          HoldPYYTDBalance@ = HoldPYYTDBalance@ + rs("PY YTD Balance")
          HoldBUDYTDBalance@ = HoldBUDYTDBalance@ + rs("BUD YTD Balance")
          rs.MoveNext
          If Err = 3021 Then Exit Do
        Loop
      End If
    End If

      rsWork("CY Period Balance") = HoldCYPeriodBalance@
      rsWork("CY Period Ratio") = IIf(TotalCYPeriodIncome@ = 0, 0, HoldCYPeriodBalance@ / TotalCYPeriodIncome@)
      rsWork("PY Period Balance") = HoldPYPeriodBalance@
      rsWork("PY Period Ratio") = IIf(TotalPYPeriodIncome@ = 0, 0, HoldPYPeriodBalance@ / TotalPYPeriodIncome@)
      rsWork("BUD Period Balance") = HoldBUDPeriodBalance@
      rsWork("BUD Period Ratio") = IIf(TotalBUDPeriodIncome@ = 0, 0, HoldBUDPeriodBalance@ / TotalBUDPeriodIncome@)
      
      rsWork("CY YTD Balance") = HoldCYYTDBalance@
      rsWork("CY YTD Ratio") = IIf(TotalCYYTDIncome@ = 0, 0, HoldCYYTDBalance@ / TotalCYYTDIncome@)
      rsWork("PY YTD Balance") = HoldPYYTDBalance@
      rsWork("PY YTD Ratio") = IIf(TotalPYYTDIncome@ = 0, 0, HoldPYYTDBalance@ / TotalPYYTDIncome@)
      rsWork("BUD YTD Balance") = HoldBUDYTDBalance@
      rsWork("BUD YTD Ratio") = IIf(TotalBUDYTDIncome@ = 0, 0, HoldBUDYTDBalance@ / TotalBUDYTDIncome@)
                                          
      rsWork("Total CY Period Income") = TotalCYPeriodIncome@
      rsWork("Total CY Period COGS") = TotalCYPeriodCOGS@
      rsWork("Total CY Period Expenses") = TotalCYPeriodExpenses@
      rsWork("Total PY Period Income") = TotalPYPeriodIncome@
      rsWork("Total PY Period COGS") = TotalPYPeriodCOGS@
      rsWork("Total PY Period Expenses") = TotalPYPeriodExpenses@
      rsWork("Total BUD Period Income") = TotalBUDPeriodIncome@
      rsWork("Total BUD Period COGS") = TotalBUDPeriodCOGS@
      rsWork("Total BUD Period Expenses") = TotalBUDPeriodExpenses@
      
      rsWork("Total CY YTD Income") = TotalCYYTDIncome@
      rsWork("Total CY YTD COGS") = TotalCYYTDCOGS@
      rsWork("Total CY YTD Expenses") = TotalCYYTDExpenses@
      rsWork("Total PY YTD Income") = TotalPYYTDIncome@
      rsWork("Total PY YTD COGS") = TotalPYYTDCOGS@
      rsWork("Total PY YTD Expenses") = TotalPYYTDExpenses@
      rsWork("Total BUD YTD Income") = TotalBUDYTDIncome@
      rsWork("Total BUD YTD COGS") = TotalBUDYTDCOGS@
      rsWork("Total BUD YTD Expenses") = TotalBUDYTDExpenses@

    rsWork.Update
      
    rsWork.MoveNext
  Loop

  rs.Close
  Set rs = Nothing
  rsWork.Close
  Set rsWork = Nothing
  rsGL.Close
  Set rsGL = Nothing
  db.Close
  Set db = Nothing
  Exit Function

BuildWorkIncomeStatement_Error:
  Call ErrorLog("Reporting Module", "BuildWorkIncomeStatement", Now, Err.Number, Err.Description, True, db)
  Resume Next

End Function

Function GetBeginningBalance(Account$) As Currency

  'On Error Resume Next

  'Calculate the beginning balance for this account
  Dim db As ADODB.Connection
  Set db = New ADODB.Connection
  db.CursorLocation = adUseClient
  db.Open gblADOProvider
  
  Dim x%
  Dim Balance@
  
  Dim rsGL As ADODB.Recordset
  rsGL.Open "GL Chart Of Accounts", db, adOpenStatic, adLockOptimistic, adCmdTable

  rsGL.Index = "PrimaryKey"
  rsGL.Seek Account$
  If rsGL.BOF And rsGL.EOF Then
    'No Account
  Else
   Balance@ = rsGL("GL COA CY Beginning Amt")
   For x% = 1 To Forms![Reporting]![Reporting Periods].Form![Start Period] - 1
     Balance@ = Balance@ + rsGL("GL COA CY Period " & Trim(CStr(x%)) & " Amt")
   Next x%
  End If
  
  GetBeginningBalance = Balance@

  rsGL.Close
  Set rsGL = Nothing
  db.Close
  Set db = Nothing

End Function

Function GetBudBalance(Account$, ThroughPeriod%) As Currency

  'Return Balance Sheet info for this account
                               
  'On Error GoTo GetBudBalance_Error
                               
  Dim db As ADODB.Connection
  Set db = New ADODB.Connection
  db.CursorLocation = adUseClient
  db.Open gblADOProvider
    
  Dim rsGL As ADODB.Recordset
  rsGL.Open "GL Chart Of Accounts", db, adOpenStatic, adLockOptimistic, adCmdTable
  
  rsGL.Index = "PrimaryKey"
  rsGL.Seek Account$
  If rsGL.BOF And rsGL.EOF Then
    GetBudBalance@ = 0
  Else
     Dim x%
     Dim TempBalance@
     TempBalance@ = rsGL("GL COA PY Beginning Amt")
     For x% = 1 To ThroughPeriod%
       TempBalance@ = TempBalance@ + rsGL("GL COA BUD Period " & Trim(CStr(x%)) & " Amt")
     Next x%
  End If

  GetBudBalance@ = TempBalance@

  Exit Function
GetBudBalance_Error:
  Call ErrorLog("Reporting Module", "GetBudBalance", Now, Err.Number, Err.Description, True, db)
  Resume Next

  rsGL.Close
  Set rsGL = Nothing
  db.Close
  Set db = Nothing
  
End Function

Function GetCYBalance(Account$, ThroughPeriod%) As Currency

  'Return Balance Sheet info for this account
                               
  'On Error GoTo GetCYBalance_Error
                               
  Dim db As ADODB.Connection
  Set db = New ADODB.Connection
  db.CursorLocation = adUseClient
  db.Open gblADOProvider
  
  Dim rsGL As ADODB.Recordset
  rsGL.Open "GL Chart Of Accounts", db, adOpenStatic, adLockOptimistic, adCmdTable
  rsGL.Index = "PrimaryKey"
  rsGL.Seek Account$
  If rsGL.BOF And rsGL.EOF Then
    GetCYBalance@ = 0
 Else
  Dim x%
  Dim TempBalance@
  TempBalance@ = rsGL("GL COA CY Beginning Amt")
    For x% = 1 To ThroughPeriod%
     TempBalance@ = TempBalance@ + rsGL("GL COA CY Period " & Trim(CStr(x%)) & " Amt")
    Next x%
  End If

  GetCYBalance@ = TempBalance@
  
  rsGL.Close
  Set rsGL = Nothing
  db.Close
  Set db = Nothing
  Exit Function

GetCYBalance_Error:
  Call ErrorLog("Reporting Module", "GetCYBalance", Now, Err.Number, Err.Description, True, db)
  Resume Next

End Function

Function GetPYBalance(Account$, ThroughPeriod%) As Currency

  'On Error GoTo GetPYBalance_Error

  'Return Balance Sheet info for this account
                               
  Dim db As ADODB.Connection
  Set db = New ADODB.Connection
  db.CursorLocation = adUseClient
  db.Open gblADOProvider
  
  Dim rsGL As ADODB.Recordset
  rsGL.Open "GL Chart Of Accounts", db, adOpenStatic, adLockOptimistic, adCmdTable
  
  rsGL.Index = "PrimaryKey"
  rsGL.Seek Account$
  
  If rsGL.BOF And rsGL.EOF Then
    GetPYBalance@ = 0
  Else
    Dim x%
    Dim TempBalance@
    TempBalance@ = rsGL("GL COA PY Beginning Amt")
    For x% = 1 To ThroughPeriod%
      TempBalance@ = TempBalance@ + rsGL("GL COA PY Period " & Trim(CStr(x%)) & " Amt")
    Next x%
  End If

  GetPYBalance@ = TempBalance@

  rsGL.Close
  Set rsGL = Nothing
  db.Close
  Set db = Nothing
  
  Exit Function
GetPYBalance_Error:
  Call ErrorLog("Reporting Module", "GetPYBalance", Now, Err.Number, Err.Description, True, db)
  Resume Next
  
End Function

Function RollBalance(Account$, TempBalance@, LEVEL%) As Currency
                               
  'On Error GoTo RollBalance_Error
                               
  Dim db As ADODB.Connection
  Set db = New ADODB.Connection
  db.CursorLocation = adUseClient
  db.Open gblADOProvider

  Dim HoldBalance@
  HoldBalance@ = TempBalance@

  If LEVEL% = 5 Then
    RollBalance@ = HoldBalance@
    Exit Function
  End If

  Dim rs As ADODB.Recordset
  rs.Open "rpt - Fin - Balance Sheet", db, adOpenStatic, adLockOptimistic, adCmdTable
  rs.MoveFirst
  rs.Find "[GL COA Account No] = '" & Account$ & "'"

  If rs("GL COA Reporting Level") > LEVEL% Then
    RollBalance@ = 0
  Else
    'On Error Resume Next
    rs.MoveNext
    If Err = 0 Then
      Do While rs("GL COA Reporting Level") > LEVEL% And Not rs.EOF
        HoldBalance@ = HoldBalance@ + rs("Balance 1")
        rs.MoveNext
        If Err = 3021 Then Exit Do
      Loop
    End If
  End If

  RollBalance@ = HoldBalance@

  rs.Close
  Set rs = Nothing
  db.Close
  Set db = Nothing
    
  Exit Function
RollBalance_Error:
  Call ErrorLog("Reporting Module", "RollBalance", Now, Err.Number, Err.Description, True, db)
  Resume Next

End Function


Function RollBudBalance(Account$, TempBalance@, LEVEL%) As Currency
                                 
  'On Error GoTo RollBudBalance_Error
                                 
  Dim db As ADODB.Connection
  Set db = New ADODB.Connection
  db.CursorLocation = adUseClient
  db.Open gblADOProvider
  
  Dim HoldBalance@
  HoldBalance@ = TempBalance@

  If LEVEL% = 5 Then
    RollBudBalance@ = HoldBalance@
    Exit Function
  End If

  Dim rs As ADODB.Recordset
  rs.Open "rpt - Fin - Balance Sheet", db, adOpenStatic, adLockOptimistic, adCmdTable
  rs.MoveFirst
  rs.Find "[GL COA Account No] = '" & Account$ & "'"

  If rs("GL COA Reporting Level") > LEVEL% Then
    RollBudBalance@ = 0
  Else
    'On Error Resume Next
    rs.MoveNext
    If Err = 0 Then
      Do While rs("GL COA Reporting Level") > LEVEL% And Not rs.EOF
        HoldBalance@ = HoldBalance@ + rs("Balance 2")
        rs.MoveNext
        If Err = 3021 Then Exit Do
      Loop
    End If
  End If

  RollBudBalance@ = HoldBalance@

  rs.Close
  Set rs = Nothing
  db.Close
  Set db = Nothing

  Exit Function
RollBudBalance_Error:
  Call ErrorLog("Reporting Module", "RollBudBalance", Now, Err.Number, Err.Description, True, db)
  Resume Next

End Function

Function RollPYBalance(Account$, TempBalance@, LEVEL%) As Currency
                                 
  'On Error GoTo RollPYBalance_Error
                                 
Dim db As ADODB.Connection
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open gblADOProvider

  Dim HoldBalance@
  HoldBalance@ = TempBalance@

  If LEVEL% = 5 Then
    RollPYBalance@ = HoldBalance@
    Exit Function
  End If

  Dim rs As ADODB.Recordset
  rs.Open "rpt - Fin - Balance Sheet", db, adOpenStatic, adLockOptimistic, adCmdTable
  rs.MoveFirst
  rs.Find "[GL COA Account No] = '" & Account$ & "'"

  If rs("GL COA Reporting Level") > LEVEL% Then
    RollPYBalance@ = 0

  Else
    'On Error Resume Next
    rs.MoveNext
    If Err = 0 Then
      Do While rs("GL COA Reporting Level") > LEVEL% And Not rs.EOF
        HoldBalance@ = HoldBalance@ + rs("Balance 2")
        rs.MoveNext
        If Err = 3021 Then Exit Do
      Loop
    End If
  End If

  RollPYBalance@ = HoldBalance@

  rs.Close
  Set rs = Nothing
  db.Close
  Set db = Nothing

  Exit Function
RollPYBalance_Error:
  Call ErrorLog("Reporting Module", "RollPYBalance", Now, Err.Number, Err.Description, True, db)
  Resume Next

End Function

