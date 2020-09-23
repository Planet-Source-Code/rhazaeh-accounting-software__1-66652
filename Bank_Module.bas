Attribute VB_Name = "Bank_Module"

Function PostDeposit(DocumentKey&) As Integer

  'On Error GoTo PostDeposit_Error

  Dim msg$
  Dim title$

  Dim db As ADODB.Connection
  Set db = New ADODB.Connection
  db.CursorLocation = adUseServer
  db.Open gblADOProvider
  
  Dim rsCompany As ADODB.Recordset
  Set rsCompany = New ADODB.Recordset
  rsCompany.Open "[SYS Company]", db, adOpenStatic, adLockOptimistic, adCmdTable
  rsCompany.MoveFirst

  Dim rsGLWorkDetail As ADODB.Recordset
  Set rsGLWorkDetail = New ADODB.Recordset
  rsGLWorkDetail.Open "[GL Work Detail]", db, adOpenStatic, adLockOptimistic, adCmdTable
  
  Dim rsBankTrans As ADODB.Recordset
  Set rsBankTrans = New ADODB.Recordset
  rsBankTrans.Open "[BANK Transaction]", db, adOpenStatic, adLockOptimistic, adCmdTable

  ' first lets get the Credit Memo
  'rsBankTrans.Index = "PrimaryKey"
  'rsBankTrans.Seek DocumentKey&
  rsBankTrans.Find "[BANK TRANS ID]=" & DocumentKey&

  If rsBankTrans("BANK TRANS Beg Balance") = True Then
    PostDeposit% = True
    Exit Function
  End If

  'Post by 1 - system date or 2 - Transaction date?
  Dim PostDate%
  PostDate% = rsCompany("SYS COM GL Post By Date")

  'Set Post Date
  Dim TranDate As Variant
  If PostDate% = 1 Then
    TranDate = DateValue(Format(Now, "Short Date"))
  Else
    TranDate = DateValue(rsBankTrans("BANK TRANS Date"))
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
    GoTo UnableToPostDepositHere
  End If

  'On Error GoTo PostDeposit_Error

  ' clear any GL Work records
  Dim cmdtemp As ADODB.Recordset
  Set cmdtemp = New ADODB.Recordset
  cmdtemp.Open "DELETE DISTINCTROW * FROM [GL Work Detail]", db, , , adCmdText
  'cmdtemp.Close
  Set cmdtemp = Nothing
    
  Dim rsGLTrans As ADODB.Recordset
  Set rsGLTrans = New ADODB.Recordset
  rsGLTrans.Open "[GL Transaction]", db, adOpenStatic, adLockOptimistic, adCmdTable

  ' write GL Transaction Header
  Dim refr$
  Dim desc$
  Dim NewNumber&
  rsGLTrans.AddNew
    
    rsGLTrans("GL TRANS Document #") = "MDEP " & rsBankTrans("BANK TRANS Ext Document No")
    
    ' gl post date
    If PostDate% = 1 Then
      rsGLTrans("GL TRANS Date") = Format(Now, "Short Date")
    Else
      rsGLTrans("GL TRANS Date") = rsBankTrans("BANK TRANS Date")
    End If
    
    rsGLTrans("GL TRANS Type") = "Misc Deposit"

    refr$ = IIf(IsNull(rsBankTrans("BANK TRANS Reference")), "", rsBankTrans("BANK TRANS Reference"))
    
    rsGLTrans("GL TRANS Reference") = refr$
    rsGLTrans("GL TRANS Amount") = rsBankTrans("BANK TRANS Amount")
    rsGLTrans("GL TRANS Posted YN") = 1
    desc$ = refr$
    If Len(Trim$(desc$)) = 0 Then
      desc$ = "MDEP " & rsBankTrans("BANK TRANS Ext Document No")
    End If
    rsGLTrans("GL TRANS Description") = desc$
    rsGLTrans("GL TRANS Source") = "MDEP " & rsBankTrans("BANK TRANS Ext Document No")
    rsGLTrans("GL TRANS System Generated") = True
  rsGLTrans.Update
    NewNumber& = rsGLTrans("GL TRANS Number")

  
  '                 Debit   Credit    Source
  '                 -----   ------    ------
  ' Bank Account      X               cboBankDep.text
  ' Combo Selection           X       txtDepositAcct$
  
  Dim rsGLTransDetail As ADODB.Recordset
  Set rsGLTransDetail = New ADODB.Recordset
  rsGLTransDetail.Open "[GL Transaction Detail]", db, adOpenStatic, adLockOptimistic, adCmdTable
  
  ' debit
  rsGLTransDetail.AddNew
    rsGLTransDetail("GL TRANSD Number") = NewNumber&
    rsGLTransDetail("GL TRANSD Account") = rsBankTrans("BANK TRANS Bank Acct 1")
    rsGLTransDetail("GL TRANSD Debit Amount") = rsBankTrans("BANK TRANS Amount")
    rsGLTransDetail("GL TRANSD Credit Amount") = 0
  rsGLTransDetail.Update

  Dim AccountPost$
  Dim CreditAmount@
  Dim DebitAmount@
  Dim Success%

  AccountPost$ = rsBankTrans("BANK TRANS Bank Acct 1")
  DebitAmount@ = rsBankTrans("BANK TRANS Amount")
  CreditAmount@ = 0
  Success% = PostCOA(AccountPost$, TranDate, DebitAmount@, CreditAmount@)
  If Success% = False Then
    MsgBox "An error occurred posting the transaction to the GL!", , "Error"
    PostDeposit% = False
    Exit Function
  End If


  ' credit
  rsGLTransDetail.AddNew
    rsGLTransDetail("GL TRANSD Number") = NewNumber&
    rsGLTransDetail("GL TRANSD Account") = rsBankTrans("BANK TRANS Bank Acct 2")
    rsGLTransDetail("GL TRANSD Debit Amount") = 0
    rsGLTransDetail("GL TRANSD Credit Amount") = rsBankTrans("BANK TRANS Amount")
  rsGLTransDetail.Update

  AccountPost$ = rsBankTrans("BANK TRANS Bank Acct 2")
  DebitAmount@ = 0
  CreditAmount@ = rsBankTrans("BANK TRANS Amount")
  Success% = PostCOA(AccountPost$, TranDate, DebitAmount@, CreditAmount@)
  If Success% = False Then
    MsgBox "An error occurred posting the transaction to the GL!", , "Error"
    PostDeposit% = False
    Exit Function
  End If
  
  PostDeposit = True
  rsCompany.Close
  Set rsCompany = Nothing
  rsGLWorkDetail.Close
  Set rsGLWorkDetail = Nothing
  rsBankTrans.Close
  Set tsBankTrans = Nothing
  rsGLTrans.Close
  Set rsGLTrans = Nothing
  rsGLTransDetail.Close
  Set rsGLTransDetail = Nothing
  db.Close
  Set db = Nothing
  
  Exit Function

UnableToPostDepositHere:
  PostDeposit = False
  rsCompany.Close
  Set rsCompany = Nothing
  rsGLWorkDetail.Close
  Set rsGLWorkDetail = Nothing
  rsBankTrans.Close
  Set tsBankTrans = Nothing
  rsGLTrans.Close
  Set rsGLTrans = Nothing
  rsGLTransDetail.Close
  Set rsGLTransDetail = Nothing
  db.Close
  Set db = Nothing
  Exit Function

PostDeposit_Error:
  Call ErrorLog("Bank Module", "PostDeposit", Now, Err.Number, Err.Description, True, db)
  PostDeposit = False
  rsCompany.Close
  Set rsCompany = Nothing
  rsGLWorkDetail.Close
  Set rsGLWorkDetail = Nothing
  rsBankTrans.Close
  Set tsBankTrans = Nothing
  rsGLTrans.Close
  Set rsGLTrans = Nothing
  rsGLTransDetail.Close
  Set rsGLTransDetail = Nothing
  db.Close
  Set db = Nothing
  Exit Function

End Function

Function PostReconciliation(db As ADODB.Connection) As Integer
Dim Currentdb As Boolean
Dim msg$
Dim title$
  
  'Dim db As ADODB.Connection
  Currentdb = False
  If db Is Nothing Then
    Set db = New ADODB.Connection
    db.CursorLocation = adUseServer
    db.Open gblADOProvider
    Currentdb = True
  End If

  'On Error GoTo PostReconciliation_Error
  
  Dim rsCompany As ADODB.Recordset
  Set rsCompany = New ADODB.Recordset
  rsCompany.Open "[SYS Company]", db, adOpenStatic, adLockOptimistic, adCmdTable
  rsCompany.MoveFirst

  Dim rsGLWorkDetail As ADODB.Recordset
  Set rsGLWorkDetail = New ADODB.Recordset
  rsGLWorkDetail.Open "[GL Work Detail]", db, adOpenStatic, adLockOptimistic, adCmdTable

  Dim rsRec As ADODB.Recordset
  Set rsRec = New ADODB.Recordset
  rsRec.Open "[Bank Reconciliation]", db, adOpenStatic, adLockOptimistic, adCmdTable
   rsRec.MoveFirst

  'Update Bank Reconciliation Detail Table from Credits & Debits
   Dim qryBankRecDetailBuild As String
   'qryBankRecDetailBuild = "INSERT INTO [BANK Reconciliation Detail] ( [BANK RECD Doc #], [BANK RECD Cleared], [BANK RECD Date], [BANK RECD Description], [BANK RECD Type], [BANK RECD Amount] )"
   'qryBankRecDetailBuild = qryBankRecDetailBuild & " SELECT [qry___Bank_Rec_Detail].[Doc #], [qry___Bank_Rec_Detail].Cleared, [qry___Bank_Rec_Detail].Date, [qry___Bank_Rec_Detail].Description, [qry___Bank_Rec_Detail].Type, [qry___Bank_Rec_Detail].Amount"
   'qryBankRecDetailBuild = qryBankRecDetailBuild & " FROM [qry___Bank_Rec_Detail]"
   qryBankRecDetailBuild = "INSERT INTO [BANK Reconciliation Detail] ( [BANK RECD Doc #], [BANK RECD Cleared], [BANK RECD Date], [BANK RECD Description], [BANK RECD Type], [BANK RECD Amount] )"
   qryBankRecDetailBuild = qryBankRecDetailBuild & " SELECT [qry - Bank Rec Detail].[Doc #], [qry - Bank Rec Detail].Cleared, [qry - Bank Rec Detail].Date, [qry - Bank Rec Detail].Description, [qry - Bank Rec Detail].Type, [qry - Bank Rec Detail].Amount"
   qryBankRecDetailBuild = qryBankRecDetailBuild & " FROM [qry - Bank Rec Detail]"
   
   Dim cmdtemp As ADODB.Recordset
   Set cmdtemp = New ADODB.Recordset
   cmdtemp.Open qryBankRecDetailBuild, db, adOpenStatic, adLockOptimistic, adCmdText
   'cmdtemp.Close
   Set cmdtemp = Nothing

  'Post by 1 - system date or 2 - Transaction date?
  Dim PostDate%
  PostDate% = rsCompany("SYS COM GL Post By Date")

  'Set Post Date
  Dim TranDate As Variant
  If PostDate% = 1 Then
    TranDate = DateValue(Format(Now, "Short Date"))
  Else
    TranDate = DateValue(IIf(IsNull(rsRec("BANK REC Cutoff Date")), Format(Now, "Short Date"), rsRec("BANK REC Cutoff Date")))
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
    GoTo UnableToPostReconciliationHere
  End If

  'On Error GoTo PostReconciliation_Error

  'clear any GL Work records
  Dim cmdtemp2 As ADODB.Recordset
  Set cmdtemp2 = New ADODB.Recordset
  cmdtemp2.Open "DELETE DISTINCTROW * FROM [GL Work Detail]", db, , , adCmdText
  'cmdtemp2.Close
  Set cmdtemp2 = Nothing
  
  Dim rsGLTrans As ADODB.Recordset
  Set rsGLTrans = New ADODB.Recordset
  rsGLTrans.Open "[GL Transaction]", db, adOpenStatic, adLockOptimistic, adCmdTable

  Dim rsGLTransDetail As ADODB.Recordset
  Set rsGLTransDetail = New ADODB.Recordset
  rsGLTransDetail.Open "[GL Transaction Detail]", db, adOpenStatic, adLockOptimistic, adCmdTable

  'Mark all cleared records as closed
  Dim rsRecDetail As ADODB.Recordset
  Set rsRecDetail = New ADODB.Recordset
  rsRecDetail.Open "[Bank Reconciliation Detail]", db, adOpenStatic, adLockOptimistic, adCmdTable

  If rsRecDetail.RecordCount < 1 Then GoTo skipclearing
  rsRecDetail.MoveFirst
  
  Dim rsPaymentHeader As ADODB.Recordset
  Set rsPaymentHeader = New ADODB.Recordset
  rsPaymentHeader.Open "SELECT * [AP Payment Header] WHERE [AP PAY ID]='" & rsRecDetail("BANK RECD Doc #") & "'", db, adOpenStatic, adLockOptimistic, adCmdText
  'rsPaymentHeader.Seek rsRecDetail("BANK RECD Description"), rsRecDetail("BANK RECD Doc #")
  
  Dim rsReceiptHeader As ADODB.Recordset
  Set rsReceiptHeader = New ADODB.Recordset
  rsReceiptHeader.Open "[AR Payment Header]", db, adOpenStatic, adLockOptimistic, adCmdTable
  
  Dim rsBankTrans As ADODB.Recordset
  Set rsBankTrans = New ADODB.Recordset
  rsBankTrans.Open "[Bank Transaction]", db, adOpenStatic, adLockOptimistic, adCmdTable

  Do While Not rsRecDetail.EOF
    Select Case rsRecDetail("BANK RECD Type")
    Case "Deposit", "Withdrawal", "Transfer From", "Transfer To", "Deposit Slip"
      'rsBankTrans.Index = "BANK TRANS Ext Document No"
      'rsBankTrans.Seek rsRecDetail("BANK RECD Doc #")
      rsBankTrans.MoveFirst
      rsBankTrans.Find "[BANK TRANS Ext Document No]='" & rsRecDetail("BANK RECD Doc #") & "'"
      If rsBankTrans.EOF Then
      Else
        rsBankTrans("BANK TRANS Cleared YN") = True
        rsBankTrans.Update
      End If
    Case "Payment", "Payroll", "Refund"
      'rsPaymentHeader.Index = "PrimaryKey"
      'rsPaymentHeader.Seek rsRecDetail("BANK RECD Description"), rsRecDetail("BANK RECD Doc #")
      MsgBox "add find function rsRecDetail![BANK RECD Description]"
      If rsPaymentHeader.EOF Then
      Else
        rsPaymentHeader("AP PAY Reconciled") = True
        rsPaymentHeader.Update
      End If
    Case "Cash Receipt"
      rsReceiptHeader.Index = "PrimaryKey"
      rsReceiptHeader.Seek rsRecDetail("BANK RECD Description"), rsRecDetail("BANK RECD Doc #")
      If rsReceiptHeader.EOF Then
      Else
        rsReceiptHeader("AR PAY Reconciled") = True
        rsReceiptHeader.Update
      End If
    End Select
    rsRecDetail.MoveNext
  Loop
skipclearing:

'------------------------------------------------------------------

  'Do Interest Earned
  '                Debit   Credit    Source
  '                -----   ------    ------
  ' Bank Expense     X
  ' Bank Account             X

  Dim NewNumber&
  Dim CreditAmount@
  Dim DebitAmount@
  Dim AccountPost$
  Dim Success%

  If (rsRec("BANK REC Interest") = 0) Then
  Else
    rsGLTrans.AddNew
      NewNumber& = rsGLTrans("GL TRANS Number")
      rsGLTrans("GL TRANS Document #") = "BANK REC " & Trim(Str(NewNumber&))
      rsGLTrans("GL TRANS Type") = "BANK"
      If PostDate% = 1 Then
        rsGLTrans("GL TRANS Date") = Format(Now, "Short Date")
      Else
        rsGLTrans("GL TRANS Date") = rsRec("BANK REC Cutoff Date")
      End If
      rsGLTrans("GL TRANS Description") = "Interest Earned" '"Bank Reconciliation"
      rsGLTrans("GL TRANS Reference") = rsRec("BANK REC Interest Acct")
      rsGLTrans("GL TRANS Source") = "Bank Reconciliation"
      rsGLTrans("GL TRANS Amount") = rsRec("BANK REC Interest")
      rsGLTrans("GL TRANS Posted YN") = 1
      rsGLTrans("GL TRANS System Generated") = True
    rsGLTrans.Update

    rsGLTransDetail.AddNew
      rsGLTransDetail("GL TRANSD Account") = rsRec("BANK REC Bank Acct")
      rsGLTransDetail("GL TRANSD Debit Amount") = rsRec("BANK REC Interest")
      rsGLTransDetail("GL TRANSD Credit Amount") = 0
      rsGLTransDetail("GL TRANSD Number") = NewNumber&
    rsGLTransDetail.Update

    AccountPost$ = rsRec("BANK REC Bank Acct")
    DebitAmount@ = rsRec("BANK REC Interest")
    CreditAmount@ = 0
    Success% = PostCOA(AccountPost$, TranDate, DebitAmount@, CreditAmount@, db)
    If Success% = False Then
      MsgBox "An error occurred posting to the GL!", , "Error"
      GoTo UnableToPostReconciliationHere
    End If

    rsGLTransDetail.AddNew
      rsGLTransDetail("GL TRANSD Account") = rsRec("BANK REC Interest Acct")
      rsGLTransDetail("GL TRANSD Debit Amount") = 0
      rsGLTransDetail("GL TRANSD Credit Amount") = rsRec("BANK REC Interest")
      rsGLTransDetail("GL TRANSD Number") = NewNumber&
    rsGLTransDetail.Update

    AccountPost$ = rsRec("BANK REC Interest Acct")
    DebitAmount@ = 0
    CreditAmount@ = rsRec("BANK REC Interest")
    Success% = PostCOA(AccountPost$, TranDate, DebitAmount@, CreditAmount@, db)
    If Success% = False Then
      MsgBox "An error occurred posting to the GL!", , "Error"
      GoTo UnableToPostReconciliationHere
    End If

  End If

'------------------------------------------------------------------

  'Do Other Charges
  '                Debit   Credit    Source
  '                -----   ------    ------
  ' Bank Expense     X
  ' Bank Account             X

  If (rsRec("BANK REC Service Charge") = 0) Then
  Else
    rsGLTrans.AddNew
      NewNumber& = rsGLTrans("GL TRANS Number")
      rsGLTrans("GL TRANS Document #") = "BANK REC " & Trim(Str(NewNumber&))
      rsGLTrans("GL TRANS Type") = "BANK"
      If PostDate% = 1 Then
        rsGLTrans("GL TRANS Date") = Format(Now, "Short Date")
      Else
        rsGLTrans("GL TRANS Date") = rsRec("BANK REC Cutoff Date")
      End If
      rsGLTrans("GL TRANS Description") = "Service Charge" '"Bank Reconciliation"
      rsGLTrans("GL TRANS Reference") = rsRec("BANK REC Service Acct")
      rsGLTrans("GL TRANS Source") = "Bank Reconciliation"
      rsGLTrans("GL TRANS Amount") = rsRec("BANK REC Service Charge")
      rsGLTrans("GL TRANS Posted YN") = 1
      rsGLTrans("GL TRANS System Generated") = True
    rsGLTrans.Update

    rsGLTransDetail.AddNew
      rsGLTransDetail("GL TRANSD Account") = rsRec("BANK REC Service Acct")
      rsGLTransDetail("GL TRANSD Debit Amount") = rsRec("BANK REC Service Charge")
      rsGLTransDetail("GL TRANSD Credit Amount") = 0
      rsGLTransDetail("GL TRANSD Number") = NewNumber&
    rsGLTransDetail.Update

    AccountPost$ = rsRec("BANK REC Service Acct")
    DebitAmount@ = rsRec("BANK REC Service Charge")
    CreditAmount@ = 0
    Success% = PostCOA(AccountPost$, TranDate, DebitAmount@, CreditAmount@, db)
    If Success% = False Then
      MsgBox "An error occurred posting to the GL!", , "Error"
      GoTo UnableToPostReconciliationHere
    End If

    rsGLTransDetail.AddNew
      rsGLTransDetail("GL TRANSD Account") = rsRec("BANK REC Bank Acct")
      rsGLTransDetail("GL TRANSD Debit Amount") = 0
      rsGLTransDetail("GL TRANSD Credit Amount") = rsRec("BANK REC Service Charge")
      rsGLTransDetail("GL TRANSD Number") = NewNumber&
    rsGLTransDetail.Update

    AccountPost$ = rsRec("BANK REC Bank Acct")
    DebitAmount@ = 0
    CreditAmount@ = rsRec("BANK REC Service Charge")
    Success% = PostCOA(AccountPost$, TranDate, DebitAmount@, CreditAmount@, db)
    If Success% = False Then
      MsgBox "An error occurred posting to the GL!", , "Error"
      GoTo UnableToPostReconciliationHere
    End If

  End If

  
  PostReconciliation = True
  
    rsCompany.Close
    Set rsCompany = Nothing
    rsGLWorkDetail.Close
    Set rsGLWorkDetail = Nothing
    rsRec.Close
    Set rsRecs = Nothing
    rsGLTrans.Close
    Set rsGLTrans = Nothing
    rsGLTransDetail.Close
    Set rsGLTransDetail = Nothing
    rsRecDetail.Close
    Set rsRecDetail = Nothing
    'rsPaymentHeader.Close
    Set rsPaymentHeader = Nothing
    'rsReceiptHeader.Close
    Set rsReceiptHeader = Nothing
    'rsBankTrans.Close
    Set rsBankTrans = Nothing
If Currentdb = True Then
    db.Close
    Set db = Nothing
End If

  Exit Function

UnableToPostReconciliationHere:
  PostReconciliation = False
    
    rsCompany.Close
    Set rsCompany = Nothing
    rsGLWorkDetail.Close
    Set rsGLWorkDetail = Nothing
    rsRec.Close
    Set rsRecs = Nothing
    rsGLTrans.Close
    Set rsGLTrans = Nothing
    rsGLTransDetail.Close
    Set rsGLTransDetail = Nothing
    rsRecDetail.Close
    Set rsRecDetail = Nothing
    rsPaymentHeader.Close
    Set rsPaymentHeader = Nothing
    rsReceiptHeader.Close
    Set rsReceiptHeader = Nothing
    rsBankTrans.Close
    Set rsBankTrans = Nothing
If Currentdb = True Then
    db.Close
    Set db = Nothing
End If
  
  Exit Function

PostReconciliation_Error:
  Call ErrorLog("Bank Module", "PostReconciliation", Now, Err.Number, Err.Description, True, db)
  PostReconciliation = False
    
    rsCompany.Close
    Set rsCompany = Nothing
    rsGLWorkDetail.Close
    Set rsGLWorkDetail = Nothing
    rsRec.Close
    Set rsRecs = Nothing
    rsGLTrans.Close
    Set rsGLTrans = Nothing
    rsGLTransDetail.Close
    Set rsGLTransDetail = Nothing
    rsRecDetail.Close
    Set rsRecDetail = Nothing
    rsPaymentHeader.Close
    Set rsPaymentHeader = Nothing
    rsReceiptHeader.Close
    Set rsReceiptHeader = Nothing
    rsBankTrans.Close
    Set rsBankTrans = Nothing
If Currentdb = True Then
    db.Close
    Set db = Nothing
End If
  
  Exit Function

End Function

Function PostTransfer(DocumentKey&) As Integer

  'On Error GoTo PostTransfer_Error

  Dim msg$
  Dim title$

  Dim db As ADODB.Connection
  Set db = New ADODB.Connection
  db.CursorLocation = adUseServer
  db.Open gblADOProvider
  Dim rsCompany As ADODB.Recordset
  Set rsCompany = New ADODB.Recordset
  rsCompany.Open "[SYS Company]", db, adOpenStatic, adLockOptimistic, adCmdTable
  rsCompany.MoveFirst

  Dim rsGLWorkDetail As ADODB.Recordset
  Set rsGLWorkDetail = New ADODB.Recordset
  rsGLWorkDetail.Open "[GL Work Detail]", db, adOpenStatic, adLockOptimistic, adCmdTable
    
  Dim rsBankTrans As ADODB.Recordset
  Set rsBankTrans = New ADODB.Recordset
  rsBankTrans.Open "[BANK Transaction]", db, adOpenStatic, adLockOptimistic, adCmdTable
  
  ' first lets get the Credit Memo
  'rsBankTrans.Index = "PrimaryKey"
  'rsBankTrans.Seek DocumentKey&
  rsBankTrans.MoveFirst
  rsBankTrans.Find "[BANK TRANS ID]=" & DocumentKey&

  'Post by 1 - system date or 2 - Transaction date?
  Dim PostDate%
  PostDate% = rsCompany("SYS COM GL Post By Date")

  'Set Post Date
  Dim TranDate As Variant
  If PostDate% = 1 Then
    TranDate = DateValue(Format(Now, "Short Date"))
  Else
    TranDate = DateValue(rsBankTrans("BANK TRANS Date"))
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
    GoTo UnableToPostTransferHere
  End If

  'On Error GoTo PostTransfer_Error

  ' clear any GL Work records
  Dim cmdtemp As ADODB.Recordset
  Set cmdtemp = New ADODB.Recordset
  cmdtemp.Open "DELETE DISTINCTROW * FROM [GL Work Detail]", db, , , adCmdText
  'cmdtemp.Close
  Set cmdtemp = Nothing
    
  Dim rsGLTrans As ADODB.Recordset
  Set rsGLTrans = New ADODB.Recordset
  rsGLTrans.Open "[GL Transaction]", db, adOpenStatic, adLockOptimistic, adCmdTable
  
  ' write GL Transaction Header
  Dim refr$
  Dim desc$
  Dim NewNumber&
  rsGLTrans.AddNew
    
    rsGLTrans("GL TRANS Document #") = "TRFR " & rsBankTrans("BANK TRANS Ext Document No")
    
    ' gl post date
    If PostDate% = 1 Then
      rsGLTrans("GL TRANS Date") = Format(Now, "Short Date")
    Else
      rsGLTrans("GL TRANS Date") = rsBankTrans("BANK TRANS Date")
    End If
    
    rsGLTrans("GL TRANS Type") = "Transfer"

    refr$ = IIf(IsNull(rsBankTrans("BANK TRANS Reference")), "", rsBankTrans("BANK TRANS Reference"))
    
    rsGLTrans("GL TRANS Reference") = refr$
    rsGLTrans("GL TRANS Amount") = rsBankTrans("BANK TRANS Amount")
    rsGLTrans("GL TRANS Posted YN") = 1
    desc$ = refr$
    If Len(Trim$(desc$)) = 0 Then
      desc$ = "TRFR " & rsBankTrans("BANK TRANS Ext Document No")
    End If
    rsGLTrans("GL TRANS Description") = desc$
    rsGLTrans("GL TRANS Source") = "TRFR " & rsBankTrans("BANK TRANS Ext Document No")
    rsGLTrans("GL TRANS System Generated") = True
  rsGLTrans.Update
  NewNumber& = rsGLTrans("GL TRANS Number")

  Dim rsGLTransDetail As ADODB.Recordset
  Set rsGLTransDetail = New ADODB.Recordset
  rsGLTransDetail.Open "[GL Transaction Detail]", db, adOpenStatic, adLockOptimistic, adCmdTable
  
  rsGLTransDetail.AddNew
    rsGLTransDetail("GL TRANSD Number") = NewNumber&
    rsGLTransDetail("GL TRANSD Account") = rsBankTrans("BANK TRANS Bank Acct 2")
    rsGLTransDetail("GL TRANSD Debit Amount") = rsBankTrans("BANK TRANS Amount")
    rsGLTransDetail("GL TRANSD Credit Amount") = 0
  rsGLTransDetail.Update

  Dim AccountPost$
  Dim DebitAmount@
  Dim CreditAmount@
  Dim Success%
  
  AccountPost$ = rsBankTrans("BANK TRANS Bank Acct 2")
  DebitAmount@ = rsBankTrans("BANK TRANS Amount")
  CreditAmount@ = 0
  Success% = PostCOA(AccountPost$, TranDate, DebitAmount@, CreditAmount@)
  If Success% = False Then
    MsgBox "An error occurred posting the transaction to the GL!", , "Error"
    PostTransfer = False
    Exit Function
  End If

  rsGLTransDetail.AddNew
    rsGLTransDetail("GL TRANSD Number") = NewNumber&
    rsGLTransDetail("GL TRANSD Account") = rsBankTrans("BANK TRANS Bank Acct 1")
    rsGLTransDetail("GL TRANSD Debit Amount") = 0
    rsGLTransDetail("GL TRANSD Credit Amount") = rsBankTrans("BANK TRANS Amount")
  rsGLTransDetail.Update

  AccountPost$ = rsBankTrans("BANK TRANS Bank Acct 1")
  DebitAmount@ = 0
  CreditAmount@ = rsBankTrans("BANK TRANS Amount")
  Success% = PostCOA(AccountPost$, TranDate, DebitAmount@, CreditAmount@)
  If Success% = False Then
    MsgBox "An error occurred posting the transaction to the GL!", , "Error"
    PostTransfer = False
    Exit Function
  End If
  
  
  'Create a cloned record of type Transfer To and reverse the accounts

  Dim rs As ADODB.Recordset
  Dim rs2 As ADODB.Recordset
  Set rs = New ADODB.Recordset
  rs.Open "[Bank Transaction]", db, adOpenStatic, adLockOptimistic, adCmdTable
  Set rs2 = New ADODB.Recordset
  rs2.Open "[Bank Transaction]", db, adOpenStatic, adLockOptimistic, adCmdTable

  'rs.Index = "PrimaryKey"
  'rs.Seek DocumentKey&
  rs.MoveFirst
  rs.Find "[BANK TRANS ID]=" & DocumentKey&

  Dim MyCounter&
  Dim MyCounter2&

  MyCounter& = DocumentKey&

  'On Error Resume Next

  Dim X%
  Dim count%
  count% = rs2.Fields.count
  rs2.AddNew
    'Add all current rs records
    'MyCounter2& = rs2("BANK TRANS ID")
    For X% = 1 To count% - 1
    If IsNull(rs(X%)) = False Then
      If rs2(X%).Type = 202 Or rs2(X%).Type = 203 Then
        rs2(X%) = rs(X%) & ""
      Else
        rs2(X%) = rs(X%)
      End If
    End If
    Next X%

    'rs2("BANK TRANS ID") = MyCounter2&
    'Create an invoice ID
    Dim rsSeek As ADODB.Recordset
    Set rsSeek = New ADODB.Recordset
    rsSeek.Open "[Bank Transaction]", db, adOpenStatic, adLockOptimistic, adCmdTable
    'rsSeek.Index = "BANK TRANS Ext Document No"
    Dim Counter%
    Counter% = 1
    Success% = False
    Do While Not Success%
      gNewInvoice$ = rs2("BANK TRANS Ext Document No") & "-" & Trim(Str(Counter%))
      'Check if this newly created document exists
      rsSeek.MoveFirst
      rsSeek.Find "[BANK TRANS Ext Document No]='" & gNewInvoice$ & "'"
      If rsSeek.EOF Then
        Success% = True
      Else
        Success% = False
        Counter% = Counter% + 1
      End If
    Loop
    rs2("BANK TRANS Ext Document No") = gNewInvoice$
    rs2("BANK TRANS Posted YN") = True
    Dim Holder$
    Holder$ = rs2("BANK TRANS Bank Acct 1")
    rs2("BANK TRANS Bank Acct 1") = rs2("BANK TRANS Bank Acct 2")
    rs2("BANK TRANS Bank Acct 2") = Holder$
    rs2("BANK TRANS Type") = "Transfer To"
  rs2.Update

  PostTransfer = True
  
  rs.Close
  Set rs = Nothing
  rs2.Close
  Set rs2 = Nothing
  rsCompany.Close
  Set rsCompany = Nothing
  rsSeek.Close
  Set rsSeek = Nothing
  rsGLTrans.Close
  Set rsGLTrans = Nothing
  rsGLTransDetail.Close
  Set rsGLTransDetail = Nothing
  rsGLWorkDetail.Close
  Set rsGLWorkDetail = Nothing
  rsBankTrans.Close
  Set rsBankTrans = Nothing
  db.Close
  Set db = Nothing
  
  Exit Function

UnableToPostTransferHere:
  PostTransfer = False
   rs.Close
  Set rs = Nothing
  rs2.Close
  Set rs2 = Nothing
  rsCompany.Close
  Set rsCompany = Nothing
  rsSeek.Close
  Set rsSeek = Nothing
  rsGLTrans.Close
  Set rsGLTrans = Nothing
  rsGLTransDetail.Close
  Set rsGLTransDetail = Nothing
  rsGLWorkDetail.Close
  Set rsGLWorkDetail = Nothing
  rsBankTrans.Close
  Set rsBankTrans = Nothing
  db.Close
  Set db = Nothing
  Exit Function

PostTransfer_Error:
  Call ErrorLog("Bank Module", "PostTransfer", Now, Err.Number, Err.Description, True, db)
  PostTransfer = False
    rs.Close
  Set rs = Nothing
  rs2.Close
  Set rs2 = Nothing
  rsCompany.Close
  Set rsCompany = Nothing
  rsSeek.Close
  Set rsSeek = Nothing
  rsGLTrans.Close
  Set rsGLTrans = Nothing
  rsGLTransDetail.Close
  Set rsGLTransDetail = Nothing
  rsGLWorkDetail.Close
  Set rsGLWorkDetail = Nothing
  rsBankTrans.Close
  Set rsBankTrans = Nothing
  db.Close
  Set db = Nothing
  Exit Function

 
End Function

Function PostWithdrawal(DocumentKey&) As Integer

  'On Error GoTo PostWithdrawal_Error

  Dim msg$
  Dim title$
  
  Dim db As ADODB.Connection
  Set db = New ADODB.Connection
  db.CursorLocation = adUseServer
  db.Open gblADOProvider
  
  Dim rsCompany As ADODB.Recordset
  Set rsCompany = New ADODB.Recordset
  rsCompany.Open "[SYS Company]", db, adOpenStatic, adLockOptimistic, adCmdTable
  
  rsCompany.MoveFirst

  Dim rsGLWorkDetail As ADODB.Recordset
  Set rsGLWorkDetail = New ADODB.Recordset
  rsGLWorkDetail.Open "[GL Work Detail]", db, adOpenStatic, adLockOptimistic, adCmdTable
  
  Dim rsBankTrans As ADODB.Recordset
  Set rsBankTrans = New ADODB.Recordset
  rsBankTrans.Open "[BANK Transaction]", db, adOpenStatic, adLockOptimistic, adCmdTable

  ' first lets get the Credit Memo
  'rsBankTrans.Index = "PrimaryKey"
  'rsBankTrans.Seek DocumentKey&
  rsBankTrans.MoveFirst
  rsBankTrans.Find "[BANK TRANS ID]=" & DocumentKey&

  'Post by 1 - system date or 2 - Transaction date?
  Dim PostDate%
  PostDate% = rsCompany("SYS COM GL Post By Date")

  'Set Post Date
  Dim TranDate As Variant
  If PostDate% = 1 Then
    TranDate = DateValue(Format(Now, "Short Date"))
  Else
    TranDate = DateValue(rsBankTrans("BANK TRANS Date"))
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
    GoTo UnableToPostWithdrawalHere
  End If

  'On Error GoTo PostWithdrawal_Error

  ' clear any GL Work records
  Dim cmdtemp As ADODB.Recordset
  Set cmdtemp = New ADODB.Recordset
  cmdtemp.Open "DELETE DISTINCTROW * FROM [GL Work Detail]", db, , , adCmdText
  'cmdtemp.Close
  Set cmdtemp = Nothing
  
  Dim rsGLTrans As ADODB.Recordset
  Set rsGLTrans = New ADODB.Recordset
  rsGLTrans.Open "[GL Transaction]", db, adOpenStatic, adLockOptimistic, adCmdTable
    
  ' write GL Transaction Header
  Dim refr$
  Dim desc$
  Dim NewNumber&
  rsGLTrans.AddNew
    
    rsGLTrans("GL TRANS Document #") = "WDRL " & rsBankTrans("BANK TRANS Ext Document No")
    
    ' gl post date
    If PostDate% = 1 Then
      rsGLTrans("GL TRANS Date") = Format(Now, "Short Date")
    Else
      rsGLTrans("GL TRANS Date") = rsBankTrans("BANK TRANS Date")
    End If
    
    rsGLTrans("GL TRANS Type") = "Withdrawal"

    refr$ = IIf(IsNull(rsBankTrans("BANK TRANS Reference")), "", rsBankTrans("BANK TRANS Reference"))
    
    rsGLTrans("GL TRANS Reference") = refr$
    rsGLTrans("GL TRANS Amount") = rsBankTrans("BANK TRANS Amount")
    rsGLTrans("GL TRANS Posted YN") = 1
    desc$ = refr$
    If Len(Trim$(desc$)) = 0 Then
      desc$ = "WDRL " & rsBankTrans("BANK TRANS Ext Document No")
    End If
    rsGLTrans("GL TRANS Description") = desc$
    rsGLTrans("GL TRANS Source") = "WDRL " & rsBankTrans("BANK TRANS Ext Document No")
    rsGLTrans("GL TRANS System Generated") = True
  rsGLTrans.Update
  NewNumber& = rsGLTrans("GL TRANS Number")
  

'                 Debit   Credit    Source
'                 -----   ------    ------
' Lookup Selection  X               txtWithdrawlAcct$
' Bank Account              X       cboBank.text


  Dim rsGLTransDetail As ADODB.Recordset
  Set rsGLTransDetail = New ADODB.Recordset
  rsGLTransDetail.Open "[GL Transaction Detail]", db, adOpenStatic, adLockOptimistic, adCmdTable
  
  rsGLTransDetail.AddNew
    rsGLTransDetail("GL TRANSD Number") = NewNumber&
    rsGLTransDetail("GL TRANSD Account") = rsBankTrans("BANK TRANS Bank Acct 1")
    rsGLTransDetail("GL TRANSD Debit Amount") = 0
    rsGLTransDetail("GL TRANSD Credit Amount") = rsBankTrans("BANK TRANS Amount")
  rsGLTransDetail.Update
      

  Dim AccountPost$
  Dim DebitAmount@
  Dim CreditAmount@
  Dim Success%
  AccountPost$ = rsBankTrans("BANK TRANS Bank Acct 1")
  DebitAmount@ = 0
  CreditAmount@ = rsBankTrans("BANK TRANS Amount")
  Success% = PostCOA(AccountPost$, TranDate, DebitAmount@, CreditAmount@)
  If Success% = False Then
    MsgBox "An error occurred posting to the GL!", , "Error"
    PostWithdrawal = False
    Exit Function
  End If

  rsGLTransDetail.AddNew
    rsGLTransDetail("GL TRANSD Number") = NewNumber&
    rsGLTransDetail("GL TRANSD Account") = rsBankTrans("BANK TRANS Bank Acct 2")
    rsGLTransDetail("GL TRANSD Debit Amount") = rsBankTrans("BANK TRANS Amount")
    rsGLTransDetail("GL TRANSD Credit Amount") = 0
  rsGLTransDetail.Update

  AccountPost$ = rsBankTrans("BANK TRANS Bank Acct 2")
  DebitAmount@ = rsBankTrans("BANK TRANS Amount")
  CreditAmount@ = 0
  Success% = PostCOA(AccountPost$, TranDate, DebitAmount@, CreditAmount@)
  If Success% = False Then
    MsgBox "An error occurred posting to the GL!", , "Error"
    PostWithdrawal = False
    Exit Function
  End If
  

  PostWithdrawal = True
  rsCompany.Close
  Set rsCompany = Nothing
  rsGLTrans.Close
  Set rsGLTrans = Nothing
  rsGLTransDetail.Close
  Set rsGLTransDetail = Nothing
  rsGLWorkDetail.Close
  Set rsGLWorkDetail = Nothing
  rsBankTrans.Close
  Set rsBankTrans = Nothing
  db.Close
  Set db = Nothing
  Exit Function

UnableToPostWithdrawalHere:
  PostWithdrawal = False
  rsCompany.Close
  Set rsCompany = Nothing
  rsGLTrans.Close
  Set rsGLTrans = Nothing
  rsGLTransDetail.Close
  Set rsGLTransDetail = Nothing
  rsGLWorkDetail.Close
  Set rsGLWorkDetail = Nothing
  rsBankTrans.Close
  Set rsBankTrans = Nothing
  db.Close
  Set db = Nothing
  Exit Function

PostWithdrawal_Error:
  Call ErrorLog("Bank Module", "PostWithdrawal", Now, Err.Number, Err.Description, True, db)
  PostWithdrawal = False
  rsCompany.Close
  Set rsCompany = Nothing
  rsGLTrans.Close
  Set rsGLTrans = Nothing
  rsGLTransDetail.Close
  Set rsGLTransDetail = Nothing
  rsGLWorkDetail.Close
  Set rsGLWorkDetail = Nothing
  rsBankTrans.Close
  Set rsBankTrans = Nothing
  db.Close
  Set db = Nothing
  Exit Function

End Function

