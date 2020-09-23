Attribute VB_Name = "Purchasing_Module"

Private Sub RedoPurchaseNumbers(db As ADODB.Connection)
On Error GoTo RedoPurchaseNumbers_Error         'try to avoid using this on error
Dim SQLstatement As String
Dim Currentdb As Boolean
  
  'Dim db As ADODB.Connection
  Currentdb = False
  If db Is Nothing Then
    Set db = New ADODB.Connection
    db.CursorLocation = adUseServer
    db.Open gblADOProvider
    Currentdb = True
  End If
  
  Dim rs As ADODB.Recordset
  'Dim rs2 As ADODB.Recordset

  'Dim cmdtemp As ADODB.Recordset
  'Set cmdtemp = New ADODB.Recordset
   db.Execute "DELETE * FROM [Payment Numbers]", , adCmdText
'  cmdtemp.Close
  'Set cmdtemp = Nothing

  Set rs = New ADODB.Recordset
  rs.Open "SELECT [AP PAY Check No], [AP PAY Bank Account] FROM [AP Payment Header]", db, adOpenKeyset, adLockReadOnly, adCmdText
  'Set rs2 = New ADODB.Recordset
  'rs2.Open "[Payment Numbers]", db, adOpenStatic, adLockOptimistic, adCmdTable
      
  If rs.RecordCount > 0 Then
   rs.MoveFirst
  Do Until rs.EOF
      'rs2.AddNew
      '  rs2("Check No") = rs("AP PAY Check No") & ""
      '  rs2("Bank") = rs("AP PAY Bank Account") & ""
      'rs2.Update
      SQLstatement = "INSERT INTO [Payment Numbers]"
      SQLstatement = SQLstatement & " ([Check No],[Bank])"
      SQLstatement = SQLstatement & " VALUES ('" & rs("AP PAY Check No") & "" & "','" & rs("AP PAY Bank Account") & "" & "')"
      db.Execute SQLstatement
      
      rs.MoveNext
    
    Loop
  End If
  
  Exit Sub
RedoPurchaseNumbers_Error:
  Call ErrorLog("Cash Payments", "RedoPurchaseNumbers", Now, Err.Number, Err.Description, True, db)
  Resume Next
    
rs.Close
Set rs = Nothing
'rs2.Close
'Set rs2 = Nothing
If Currentdb = True Then
    db.Close
    Set db = Nothing
End If
    
End Sub

Function BuildAgedPayables() As Integer

  'On Error GoTo BuildAgedPayables_Error

  Dim db As ADODB.Connection
  Set db = New ADODB.Connection
  db.CursorLocation = adUseServer
  db.Open gblADOProvider
  
  Dim rsWork As ADODB.Recordset

  Dim cmdtemp As ADODB.Recordset
  Set cmdtemp = New ADODB.Recordset
  cmdtemp.Open "DELETE * FROM [Print Aged Payables Work]", db, , , adCmdText
  cmdtemp.Close
  Set cmdtemp = Nothing

   Set rsWork = New ADODB.Recordset
   rsWork.Open "[Print Aged Payables Work]", db, adOpenStatic, adLockOptimistic, adCmdTable
                              
  Dim rsPayments As ADODB.Recordset
  Dim VendorID$

  Dim rsCompany As ADODB.Recordset
  Set rsCompany = New ADODB.Recordset
   rsCompany.Open "[SYS Company]", db, adOpenStatic, adLockOptimistic, adCmdTable

  Dim Period1%
  Dim Period2%
  Dim Period3%
  Dim AgeBy%

  rsCompany.MoveFirst
  Period1% = rsCompany("SYS COM Purchase Period 1")
  Period2% = rsCompany("SYS COM Purchase Period 2")
  Period3% = rsCompany("SYS COM Purchase Period 3")
  AgeBy% = IIf(IsNull(rsCompany("SYS COM Purchase Age Invoices By")), 1, rsCompany("SYS COM Sales Age Invoices By"))
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
  Dim rsAPPurchase As ADODB.Recordset
  Dim rsAPVendor As ADODB.Recordset
  Set rsAPVendor = New ADODB.Recordset
  rsAPVendor.Open "AP Vendor", db, adOpenStatic, adLockOptimistic, adCmdTable
  
  Dim Order%
  Order% = 0
  
  db.BeginTrans
  
  rsAPVendor.MoveFirst
  Do While Not rsAPVendor.EOF
    VendorID$ = rsAPVendor("AP VEN ID")
    Set rsAPPurchase = New ADODB.Recordset
    rsAPPurchase.Open "SELECT * FROM [AP Purchase] WHERE [AP PO Vendor ID] = '" & VendorID$ & "'AND [AP PO Document Type] in ('Receiving','Beginning Balance','Voucher') AND [AP PO Posted YN] = TRUE ORDER BY [AP PO Date] Asc", db, adOpenStatic, adLockOptimistic, adCmdText
    
    'On Error Resume Next
    rsAPPurchase.MoveFirst
    Do While Not rsAPPurchase.EOF
      rsWork.AddNew
        rsWork("Vendor ID") = VendorID$
        rsWork("Order") = Order%
        rsWork("Transaction Type") = rsAPPurchase("AP PO Document Type")
        rsWork("Transaction ID") = rsAPPurchase("AP PO Ext Document No")
        rsWork("Transaction Description") = rsAPPurchase("AP PO Description")
        rsWork("Applied To") = ""
  
        'What bucket do we use?
        'Get a date to age by
        If (AgeBy% = 1) Then 'Use Invoice Date
          TransDate = IIf(IsNull(rsAPPurchase("AP PO Date")), Format("1/1/" & Format(Now, "yyyy"), "Short Date"), rsAPPurchase("AP PO Date"))
        Else                 'Use Due Date
          TransDate = IIf(IsNull(rsAPPurchase("AP PO Due Date")), Format("1/1/" & Format(Now, "yyyy"), "Short Date"), rsAPPurchase("AP PO Due Date"))
        End If
        
        rsWork("Transaction Date") = TransDate
  
        TransAmount@ = rsAPPurchase("AP PO Total Amount")
  
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
  
      'Now load all payments to this invoice
      Err = 0
      Set rsPayments = New ADODB.Recordset
      rsPayments.Open "SELECT * FROM [qryVendorPayments] where [AP Payment Invoice Cross Reference].[AP CROSS Payed ID] = " & rsAPPurchase("AP PO Document No") & " AND [AP PAY Posted YN] = TRUE", db, adOpenStatic, adLockOptimistic, adCmdText
  
      'If Err <> 0 Then GoTo SkipAgedPayments
      If rsPayments.RecordCount = 0 Then GoTo SkipAgedPayments
  
      rsPayments.MoveFirst
      'If Err <> 0 Then GoTo SkipAgedPayments
      
      Do While Not rsPayments.EOF
        If rsPayments("AP CROSS Applied Amount") >= 0.01 Then
          rsWork.AddNew
            rsWork("Vendor ID") = VendorID$
            rsWork("Order") = Order%
            rsWork("Transaction Date") = rsPayments("AP PAY Transaction Date")
            rsWork("Transaction Type") = rsPayments("AP PAY Type")
            rsWork("Transaction ID") = rsPayments("AP PAY Check No")
            rsWork("Transaction Description") = "Applied to " & rsAPPurchase("AP PO Ext Document No")
            rsWork("Applied To") = rsAPPurchase("AP PO Ext Document No")

            Select Case CurrentPeriod%
            Case 1
              rsWork("Period 1") = rsPayments("AP CROSS Applied Amount") * -1
            Case 2
              rsWork("Period 2") = rsPayments("AP CROSS Applied Amount") * -1
            Case 3
              rsWork("Period 3") = rsPayments("AP CROSS Applied Amount") * -1
            Case 4
              rsWork("Period 4") = rsPayments("AP CROSS Applied Amount") * -1
            End Select
  
            rsWork("Balance") = rsPayments("AP CROSS Applied Amount") * -1
          rsWork.Update
          Order% = Order% + 1
    
          TotalAmount@ = TotalAmount@ - rsPayments("AP CROSS Applied Amount")
        End If
  
        If rsPayments("AP CROSS Discount Taken") >= 0.01 Then
          rsWork.AddNew
            rsWork("Vendor ID") = VendorID$
            rsWork("Order") = Order%
            rsWork("Transaction Date") = rsPayments("AP PAY Transaction Date")
            rsWork("Transaction Type") = "Discount"
            rsWork("Transaction ID") = rsPayments("AP PAY Check No")
            rsWork("Transaction Description") = "Applied to " & rsAPPurchase("AP PO Ext Document No")
            rsWork("Applied To") = rsAPPurchase("AP PO Ext Document No")
            
            Select Case CurrentPeriod%
            Case 1
              rsWork("Period 1") = rsPayments("AP CROSS Discount Taken") * -1
            Case 2
              rsWork("Period 2") = rsPayments("AP CROSS Discount Taken") * -1
            Case 3
              rsWork("Period 3") = rsPayments("AP CROSS Discount Taken") * -1
            Case 4
              rsWork("Period 4") = rsPayments("AP CROSS Discount Taken") * -1
            End Select
            
            rsWork("Balance") = rsPayments("AP CROSS Discount Taken") * -1
          rsWork.Update
          
          Order% = Order% + 1
  
          TotalAmount@ = TotalAmount@ - rsPayments("AP CROSS Discount Taken")
        End If
  
        rsPayments.MoveNext
        'rsAPPurchase.Close
      Loop
      rsPayments.Close
SkipAgedPayments:
      rsAPPurchase.MoveNext
    Loop
  
    'Now add payments to this vendor that are not applied
    Err = 0
    Set rsPayments = New ADODB.Recordset
    rsPayments.Open "SELECT * FROM [qryVendorPayments2] where [AP Payment Header].[AP PAY Vendor No] = '" & VendorID$ & "' AND [AP Payment Header].[AP PAY Unapplied Amount] > 0 AND [AP Payment Header].[AP PAY Posted YN] = true", db, adOpenStatic, adLockOptimistic, adCmdText
  
    'If Err <> 0 Then GoTo SkipAgedPayments2
    If rsPayments.RecordCount = 0 Then GoTo SkipAgedPayments2
    rsPayments.MoveFirst
    'If Err <> 0 Then GoTo SkipAgedPayments2
    Do While Not rsPayments.EOF
      rsWork.AddNew
        rsWork("Vendor ID") = VendorID$
        rsWork("Order") = Order%
        rsWork("Transaction Date") = rsPayments("AP PAY Transaction Date")
        rsWork("Transaction Type") = rsPayments("AP PAY Type")
        rsWork("Transaction ID") = rsPayments("AP PAY Check No")
        rsWork("Transaction Description") = "Unapplied"
        rsWork("Applied To") = ""
        
        'Should I be aging this???
        Days& = DateDiff("d", rsWork("Transaction Date"), Now)
        Select Case Days&
        Case Is < 0
          rsWork("Period 1") = rsPayments("AP PAY UnApplied Amount") * -1
        Case 0 To Period1%
          rsWork("Period 1") = rsPayments("AP PAY UnApplied Amount") * -1
        Case Period1% To Period2%
          rsWork("Period 2") = rsPayments("AP PAY UnApplied Amount") * -1
        Case Period2% To Period3%
          rsWork("Period 3") = rsPayments("AP PAY UnApplied Amount") * -1
        Case Else
          rsWork("Period 4") = rsPayments("AP PAY UnApplied Amount") * -1
        End Select

         Balance@ = Balance@ - rsPayments("AP PAY UnApplied Amount")
        
        rsWork("Balance") = rsPayments("AP PAY UnApplied Amount") * -1
      
      rsWork.Update
      Order% = Order% + 1
    
      rsPayments.MoveNext
    Loop

SkipAgedPayments2:
    rsAPVendor.MoveNext
  Loop
      
  db.CommitTrans

  rsWork.Close
  Set rsWork = Nothing
  rsPayments.Close
  Set rsPayments = Nothing
  rsCompany.Close
  Set rsCompay = Nothing
  rsAPVendor.Close
  Set rsAPVendor = Nothing
  rsAPPurchase.Close
  Set rsAPPurchase = Nothing
  db.Close
  Set db = Nothing
  
  Exit Function
BuildAgedPayables_Error:
  Call ErrorLog("Purchase Module", "BuildAgedPayables", Now, Err.Number, Err.Description, True, db)
  Resume Next

End Function

Function ClonePayment(DocumentKey&, db As ADODB.Connection) As Integer

  'On Error GoTo ClonePayment_Error

  'If db Is Nothing Then
  'Dim db As ADODB.Connection
  '  Set db = New ADODB.Connection
  '  db.CursorLocation = adUseServer
  '  db.Open gblADOProvider
  'End If
  
  Dim rs As ADODB.Recordset
  Dim rs2 As ADODB.Recordset

  'Dim rsRecur As ADODB.Recordset
  'Set rsRecur = New ADODB.Recordset
  'rsRecur.Open "[SYS Recurred]", db, adOpenStatic, adLockOptimistic, adCmdTable
  
  Set rs = New ADODB.Recordset
  rs.Open "SELECT * FROM [AP Payment Header] FROM [AP PAY ID]=" & DocumentKey&, db, adOpenKeyset, adLockOptimistic, adCmdText
  Set rs2 = New ADODB.Recordset
  'rs2.Open "[AP Payment Header]", db, adOpenStatic, adLockOptimistic, adCmdTable
  rs2.Open "SELECT * FROM [AP Payment Header] FROM [AP PAY ID]=" & DocumentKey&, db, adOpenKeyset, adLockOptimistic, adCmdText

  'rs.Index = "AP PAY ID"
  'rs.Seek DocumentKey&
  'rs.MoveFirst
  'rs.Find "[AP PAY ID]=" & DocumentKey&   '<<<---use select statement

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
    'MyCounter2& = rs2("AP PAY ID")
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

    'rs2("AP PAY ID") = MyCounter2&
    'Rename Check #
    'If AskForID% = True Then
    '  gNewInvoice$ = InputBox("Enter new check #")
    'Else
      'Create an invoice ID
      'Get a check number
    '  Dim Account$
    '  Account$ = rs("AP PAY Bank Account")
    '  Dim rsNumber As ADODB.Recordset
    '  rsNumber.Open "SELECT * FROM [Payment Numbers] WHERE [Bank] = '" & Account$ & "' ORDER BY Val([Check No]) DESC", db, adOpenStatic, adLockOptimistic, adCmdText
    '  rsNumber.MoveFirst
    '  gNewInvoice$ = Trim(CStr(Val(rsNumber("Check No")) + 1))
    'End If
    'If gNewInvoice$ = "" Then
    '  db.RollbackTrans
    '  ClonePayment% = 1
    '  Exit Function
    'End If
    rs2("AP PAY Transaction Date") = CDate(FormatDate(Date))
    rs2("AP PAY Check No") = CheckNumberCHQ("READ", db, rs![AP PAY Bank Account])
    rs2("AP PAY Recurring YN") = False
    rs2("AP PAY Posted YN") = False
    'rsRecur.AddNew
    '  rsRecur("Document Type") = "Cash Payment"
    '  rsRecur("Document Number") = rs2("AP PAY Check No")
    '  rsRecur("Reference") = rs2("AP PAY Vendor No")
    '  rsRecur("Amount") = rs2("AP PAY Amount")
    'rsRecur.Update
  rs2.Update

  'Call RedoPurchaseNumbers
  
  db.CommitTrans
  ClonePayment% = True
   
   rs.Close
   Set rs = Nothing
   rs2.Close
   Set rs2 = Nothing
   'rsNumber.Close
   'Set rsNumber = Nothing
   'rsRecur.Close
   'Set rsRecur = Nothing
  'db.Close
  'Set db = Nothing
  Exit Function

CopyPaymentFailed:
  db.RollbackTrans
  ClonePayment% = False
   rs.Close
   Set rs = Nothing
   rs2.Close
   Set rs2 = Nothing
   'rsNumber.Close
   'Set rsNumber = Nothing
   'rsRecur.Close
   'Set rsRecur = Nothing
  'db.Close
  'Set db = Nothing
  Exit Function
  
ClonePayment_Error:
  Call ErrorLog("Purchase Module", "ClonePayment", Now, Err.Number, Err.Description, True, db)
  Resume Next

End Function

Function ClonePurchase(DocumentKey&, db As ADODB.Connection) As Integer
  
  'On Error GoTo ClonePurchase_Error
  
  'Dim db As ADODB.Connection
  'Set db = New ADODB.Connection
  'db.CursorLocation = adUseServer
  'db.Open gblADOProvider
  
  Dim PrevPONo As String
  Dim rs As ADODB.Recordset
  Dim rs2 As ADODB.Recordset

  Dim rsDetail As ADODB.Recordset
  Dim rsDetail2 As ADODB.Recordset

  'Dim rsRecur As ADODB.Recordset
  'Set rsRecur = New ADODB.Recordset
  'rsRecur.Open "[SYS Recurred]", db, adOpenStatic, adLockOptimistic, adCmdTable
  
  Set rs = New ADODB.Recordset
  'called the table using the sqlstatement is faster with WHERE, no need to use seek anymore--->>it will be done later
  rs.Open "select * FROM [AP Purchase] WHERE [AP PO Document No]=" & DocumentKey&, db, adOpenKeyset, adLockOptimistic, adCmdText
  PrevPONo = rs![AP PO Ext Document No]
  
  Set rs2 = New ADODB.Recordset
  rs2.Open "[AP Purchase]", db, adOpenStatic, adLockOptimistic, adCmdTable
  
  'only on direct table we can use the seek --->>adCmdTableDirect
  'rs.Index = "PrimaryKey"
  'rs.MoveFirst
  'rs.Find "[AP PO Document No]=" & DocumentKey&
  

  Dim MyCounter&
  Dim MyCounter2&

  MyCounter& = DocumentKey&

  'On Error GoTo CopyFailed
  db.BeginTrans

  Dim X%
  Dim count%
  count% = rs2.Fields.count
  rs2.AddNew
    'Add all current rs records
    'MyCounter2& = rs2("AP PO Document No")
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
    'rs2("AP PO Document No") = MyCounter2&
    'Rename Ext Document #
    'If AskForInvoice% = True Then
    '  gNewInvoice$ = InputBox("Enter new document #")
    'Else
      'Create an invoice ID
      
      'Dim rsSeek As ADODB.Recordset
      'Set rsSeek = New ADODB.Recordset
      'rsSeek.Open "SELECT [AP PO Ext Document No] FROM [AP Purchase]", db, adOpenKeyset, adLockOptimistic, adCmdText
      'rsSeek.Index = "Ext Document No"
      'Dim Counter%
      'Counter% = 1
      'Dim Success%
      'Success% = False
      'Do While Not Success%
      '  gNewInvoice$ = rs2("AP PO Ext Document No") & "-" & Trim(Str(Counter%))
        'Check if this newly created document exists
      '  rsSeek.MoveFirst
      '  rsSeek.Find "[AP PO Ext Document No]='" & gNewInvoice$ & "'"
        'If rsSeek.BOF And rsSeek.EOF Then
      '  If rsSeek.EOF Then
      '    Success% = True
      '  Else
      '    Success% = False
      '    Counter% = Counter% + 1
          'MsgBox rsSeek![AP PO Ext Document No]
      '  End If
      'Loop
      'rsSeek.Close
      'Set rsSeek = Nothing
      
    'End If
    'If gNewInvoice$ = "" Then
    '  db.RollbackTrans
    '  ClonePurchase% = 1
    '  Exit Function
    'End If
    
    If rs("AP PO Document Type") = "PO" Then
      rs2("AP PO Document Type") = "Receiving"
    End If
    rs2("AP PO Date") = CDate(FormatDate(Date))
    rs2("AP PO Ext Document No") = "ConvertPURC" & AppLoginName
    rs2("AP PO Existing PO Number") = PrevPONo
    rs2![AP PO Status] = "Open"
    rs2![AP PO Saved YN] = 0
    rs2![AP PO Subtotal] = 0
    rs2![AP PO Total Amount] = 0
    rs2("AP PO Recurring YN") = False
    rs2("AP PO Posted YN") = False
    rs2("AP PO Amount Paid") = 0
    rs2("AP PO Status") = "Open"
    'rsRecur.AddNew
    '  rsRecur("Document Type") = rs2("AP PO Document Type")
    '  rsRecur("Document Number") = rs2("AP PO Ext Document No")
    '  rsRecur("Reference") = rs2("AP PO Vendor ID")
    '  rsRecur("Amount") = rs2("AP PO Total Amount")
    'rsRecur.Update
    
  rs2.Update
  rs2.Close
  Set rs2 = Nothing
  
  Dim DetailCounter&

  Set rs2 = New ADODB.Recordset
  rs2.Open "SELECT [AP PO Document No],[AP PO Ext Document No] FROM [AP Purchase] where [AP PO Ext Document No] ='ConvertPURC" & AppLoginName & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
    MyCounter2& = rs2![AP PO Document No]
    rs2![AP PO Ext Document No] = "[" & AppLoginName & Format(Now, "MMdd") & Right(Format(MyCounter2&, "0000"), 4) & "]"
  rs2.Update
  rs2.Close
  Set rs2 = Nothing
  
  rs![AP PO Description] = "This PO has been converted " & "[" & AppLoginName & Format(Now, "MMdd") & Right(Format(MyCounter2&, "0000"), 4) & "]"
  rs![AP PO Status] = "Converted"
  rs![AP PO Saved YN] = 0
  rs.Update
  rs.Close
  Set rs = Nothing

  
  Set rsDetail = New ADODB.Recordset
  rsDetail.Open "SELECT * FROM [AP Purchase Detail] where [AP POD Document No] = " & MyCounter&, db, adOpenKeyset, adLockOptimistic, adCmdText
  'On Error Resume Next
  'Err = 0
  'rsDetail.MoveLast
  If rsDetail.RecordCount = 0 Then
    'No Detail
  Else
     rsDetail.MoveFirst
    'On Error GoTo CopyFailed
    'Create new detail record
    Set rsDetail2 = New ADODB.Recordset
    rsDetail2.Open "SELECT * FROM [AP Purchase Detail]", db, adOpenKeyset, adLockOptimistic, adCmdText
    Do While Not rsDetail.EOF
      count% = rsDetail.Fields.count
'again:
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
        rsDetail2("AP POD Document No") = MyCounter2&
        'Add rest of detail records
      rsDetail2.Update
      'rsDetail2.CancelUpdate
      'GoTo again
      rsDetail.MoveNext
    Loop

    rsDetail2.Close
    Set rsDetail2 = Nothing
  End If
  
    rsDetail.Close
    Set rsDetail = Nothing

SkipDetail:
  
  db.CommitTrans
  ClonePurchase% = True
'   rs.Close
'   Set rs = Nothing
'   rs2.Close
'   Set rs2 = Nothing
   'rsRecur.Close
   'Set rsRecur = Nothing
'   rsSeek.Close
   'Set rsSeek = Nothing
   'rsDetail.Close
   'Set rsDetail = Nothing
   'rsDetail2.Close
   'Set rsDetail2 = Nothing
   'db.Close
   'Set db = Nothing
  Exit Function

CopyFailed:
  db.RollbackTrans
  ClonePurchase% = False
'   rs.Close
'   Set rs = Nothing
'   rs2.Close
'   Set rs2 = Nothing
   'rsRecur.Close
   'Set rsRecur = Nothing
   'rsSeek.Close
   'Set rsSeek = Nothing
   'rsDetail.Close
   'Set rsDetail = Nothing
   'rsDetail2.Close
   'Set rsDetail2 = Nothing
   'db.Close
   'Set db = Nothing
  Exit Function

ClonePurchase_Error:
  Call ErrorLog("Purchase Module", "ClonePurchase", Now, Err.Number, Err.Description, True, db)
  Resume Next

End Function

Function GetAPInvoiceDiscount(InvoiceID&, vntInDate As String)

  'On Error GoTo GetAPInvoiceDiscount_Error

  Dim ReferNo&
  Dim vntInvoiceDate As Variant
  Dim PaymentTerms$
  Dim Discount#
  Dim DiscountDays#
  Dim Diff%
  Dim DiscountAmount@
  Dim vntDate As String

  Dim db As ADODB.Connection
  Set db = New ADODB.Connection
  db.CursorLocation = adUseServer
  db.Open gblADOProvider

  vntDate = vntInDate

  Dim rsPurchase As ADODB.Recordset
  Set rsPurchase = New ADODB.Recordset
  rsPurchase.Open "[AP Purchase]", db, adOpenStatic, adLockOptimistic, adCmdTable

  Dim rsTerms As ADODB.Recordset
  Set rsTerms = New ADODB.Recordset
  rsTerms.Open "[List Payment Terms]", db, adOpenStatic, adLockOptimistic, adCmdTable

  'rsPurchase.Index = "PrimaryKey"
  'rsPurchase.Seek InvoiceID&
  rsPurchase.MoveFirst
  rsPurchase.Find "[AP PO Document No]='" & InvoiceID& & "'"
  'If rsPurchase.BOF And rsPurchase.EOF Then
  If rsPurchase.RecordCount = 0 Then
    GetAPInvoiceDiscount = 0#
  Else
    If IsNull(vntDate) Then vntDate = DateValue(Format(Now, "Short Date"))
    If IsDate(vntDate) Then
    Else
      vntDate = DateValue(Format(Now, "Short Date"))
    End If
    vntInvoiceDate = rsPurchase("AP PO Date")
    PaymentTerms$ = IIf(IsNull(rsPurchase("AP PO Payment Terms")), "", rsPurchase("AP PO Payment Terms"))
    'rsTerms.Index = "PrimaryKey"
    'rsTerms.Seek  PaymentTerms$
    rsTerms.MoveFirst
    rsTerms.Find "[LIST PAY Description]='" & PaymentTerms$ & "'"
    'If rsTerms.BOF And rsTerms.EOF Then
    If rsTerms.RecordCount = 0 Then
      GetAPInvoiceDiscount = 0#
    Else
      Discount# = rsTerms("LIST PAY Discount")
      If Discount# = 0 Then
        GetAPInvoiceDiscount = 0#
      Else
        Discount# = Discount# / 100
      End If
      DiscountDays# = rsTerms("LIST PAY Discount Days")
      Diff% = DateDiff("d", vntInvoiceDate, vntDate)
      If Diff% <= DiscountDays# Then
        DiscountAmount@ = rsPurchase("AP PO Total Amount") * Discount#
        GetAPInvoiceDiscount = Round(CDbl(DiscountAmount@))
      Else
        GetAPInvoiceDiscount = 0#
      End If
    End If
  End If

  rsPurchase.Close
  Set rsPurchase = Nothing
  rsTerms.Close
  Set rsTerms = Nothing
  db.Close
  Set db = Nothing
  
  Exit Function
GetAPInvoiceDiscount_Error:
  Call ErrorLog("Purchase Module", "GetAPInvoiceDiscount", Now, Err.Number, Err.Description, True, db)
  Resume Next

End Function

Sub MonthEndPurchases(db As ADODB.Connection)
On Error GoTo MonthEndPurchases_Error
  'Dim db As ADODB.Connection
  'Set db = New ADODB.Connection
  'db.CursorLocation = adUseClient
  'db.Open gblADOProvider

  'On Error GoTo MonthEndPurchases_Error
  
  db.BeginTrans
  
  'Dim rs As ADODB.Recordset
  ' Set rs = New ADODB.Recordset
  ' rs.Open "SELECT * from [AP Purchase] where [AP PO Cleared YN] = 0 and [AP PO Balance Due] <0.01", db, adOpenStatic, adLockOptimistic, adCmdText
  db.Execute "UPDATE [AP Purchase] SET [AP OPEN Cleared]=True WHERE [AP PO Cleared YN] = 0 and [AP PO Balance Due] <0.01", , adCmdText
  'On Error Resume Next
  'Err = 0
  'If Err = 0 Then
  'If rs.RecordCount = 0 Then
  '  rs.MoveFirst
  '  Do While Not rs.EOF
      'Mark record as cleared
  '      rs("AP OPEN Cleared") = True
  '    rs.Update
  '    rs.MoveNext
  '  Loop
  'End If
  
  'Dim rs2 As ADODB.Recordset
  'Set rs2 = New ADODB.Recordset
  'rs2.Open "SELECT * from [AP PAYMENT Header] where [AP PAY Cleared] = 0 and [AP PAY UnApplied Amount] <0.01", db, adOpenStatic, adLockOptimistic, adCmdText
  db.Execute "UPDATE [AP PAYMENT Header] SET [AP PAY Cleared]=True where [AP PAY Cleared] = 0 and [AP PAY UnApplied Amount] <0.01", , adCmdText
  'On Error Resume Next
  'Err = 0
  'If rs2.RecordCount = 0 Then
  '  rs2.MoveFirst
  '  Do While Not rs2.EOF
      'Mark record as cleared
  '      rs2("AP PAY Cleared") = True
  '    rs2.Update
  '    rs2.MoveNext
  '  Loop
  'End If

  db.CommitTrans

  'rs.Close
  'Set rs = Nothing
  'rs2.Close
  'Set rs2 = Nothing
  'db.Close
  'Set db = Nothing
  Exit Sub

MonthEndPurchases_Error:
  Call ErrorLog("Purchase Module", "MonthEndPurchases", Now, Err.Number, Err.Description, True, db)
  db.RollbackTrans
  'rs.Close
  'Set rs = Nothing
  'rs2.Close
  'Set rs2 = Nothing
  'db.Close
  'Set db = Nothing
  Exit Sub

End Sub

Function PostPayments(VendorID$, CheckNo$, BankID$, Optional db As ADODB.Connection) As Integer
Dim Currentdb As Boolean
  
  'Dim db As ADODB.Connection
  Currentdb = False
  If db Is Nothing Then
    Set db = New ADODB.Connection
    db.CursorLocation = adUseServer
    db.Open gblADOProvider
    Currentdb = True
  End If

  'On Error GoTo PostPayments_Error
  Dim CurrentBalance@
  Dim msg$
  Dim title$

  Dim rsPayment As ADODB.Recordset
'again:   'AP PAY Posted YN
  Set rsPayment = New ADODB.Recordset
  'rsPayment.Open "SELECT * FROM [AP Payment Header] WHERE [AP PAY Check No]='" & CheckNo$ & "' AND [AP PAY Bank Account]='" & BankID$ & "'", db, adOpenStatic, adLockOptimistic, adCmdText
  rsPayment.Open "SELECT [AP PAY Transaction Date], [AP PAY Amount], [AP PAY ID], " & _
  "[AP PAY Bank Account],[AP PAY Posted YN] FROM [AP Payment Header] WHERE [AP PAY Check No]='" & _
  CheckNo$ & "' AND [AP PAY Vendor No]='" & VendorID$ & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
  'rsPayment.Index = "BankKey"
  'rsPayment.Seek CheckNo$, BankID$
  
  ' don't post if already posted
  'GoTo again
  If rsPayment("AP PAY Posted YN") = True Then
     GoTo AllreadyPosted:
  End If
  
  Dim rsCompany As ADODB.Recordset
  Set rsCompany = New ADODB.Recordset
  'rsCompany.Open "[SYS Company]", db, adOpenStatic, adLockOptimistic, adCmdTable
  rsCompany.Open "SELECT  [SYS COM Purchase AP Acct],[SYS COM Purchase Discount Acct]," & _
  "[SYS COM GL Post By Date] FROM [SYS Company]", db, adOpenKeyset, adLockOptimistic, admcdtext
  rsCompany.MoveFirst
  '[SYS COM Purchase Inventory Acct],[SYS COM Purchase Freight Acct]
  '[SYS COM Purchase Misc Acct]
  ' clear any GL Work records
  'Dim cmdtemp As ADODB.Recordset
  'Set cmdtemp = New ADODB.Recordset
  'cmdtemp.Open "DELETE DISTINCTROW * FROM [GL Work Detail]", db, , , adCmdText
  'cmdtemp.Close
  'Set cmdtemp = Nothing
  
  db.Execute "DELETE DISTINCTROW * FROM [GL Work Detail]", , adCmdText
  
  'Dim rsGLWorkDetail As ADODB.Recordset
  'Set rsGLWorkDetail = New ADODB.Recordset
  'rsGLWorkDetail.Open "[GL Work Detail]", db, adOpenStatic, adLockOptimistic, adCmdTable

  'Post by 1 - system date or 2 - Transaction date?
  Dim PostDate%
  PostDate% = rsCompany("SYS COM GL Post By Date")

  'Set Invoice Type
  Dim TranDate As Variant

  'Set Post Date
  If PostDate% = 1 Then
    TranDate = DateValue(Format(Now, "Short Date"))
  Else
    TranDate = DateValue(rsPayment("AP PAY Transaction Date"))
  End If
  
  'Verify period can be posted to
  'Send TranDate
  'Return PeriodToPost and PeriodClosed
  Dim PeriodToPost%
  Dim PeriodClosed%
  
  Call VerifyPeriod(TranDate, PeriodToPost%, PeriodClosed%, db)
  
  If PeriodClosed% = True Then
    MsgBox "Unable to post transaction to a closed period.", , "Post Payment Error"
    PostPayments% = False
    GoTo AllreadyClosed:
  End If
 
  'Dim rsGLTrans As ADODB.Recordset
  'Set rsGLTrans = New ADODB.Recordset
  'rsGLTrans.Open "[GL Transaction]", db, adOpenStatic, adLockOptimistic, adCmdTable

  ' write GL Transaction Header
  Dim refr$
  Dim desc$
  Dim NewNumber&
  
  Dim rsVendor As ADODB.Recordset
  Set rsVendor = New ADODB.Recordset
  
  'rsVendor.Open "[AP Vendor]", db, adOpenStatic, adLockOptimistic, adCmdTable
  rsVendor.Open "SELECT [AP VEN ID], [AP VEN Name], [AP VEN Payments YTD], " & _
  "[AP VEN Payments Lifetime], [AP VEN Financial Period 1] FROM [AP Vendor] " & _
  "WHERE [AP VEN ID]='" & VendorID$ & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
  'rsVendor.Index = "PrimaryKey"
  'rsVendor.Seek VendorID$
  '[AP VEN Payment Number Lifetime],[AP VEN Payment Number YTD],[AP VEN Purchase YTD]
  '[AP VEN Purchase Lifetime],[AP VEN Purchase Number Lifetime],[AP VEN Purchase Number YTD]
  '
  'rsVendor.MoveFirst
  'rsVendor.Find "[AP VEN ID]='" & VendorID$ & "'"

  'rsGLTrans.AddNew
    
  '  rsGLTrans("GL TRANS Document #") = "CASH PAY " & CheckNo$ & "-" & rsPayment("AP PAY Bank Account")
  '  rsGLTrans("GL TRANS Type") = "Cash Payment"
    
    Dim SQLstatement As String
    ' gl post date
      SQLstatement = "INSERT INTO [GL Transaction]"
      SQLstatement = SQLstatement & " ([GL TRANS Document #],[GL TRANS Type],[GL TRANS Date],"
      SQLstatement = SQLstatement & " [GL TRANS Reference],[GL TRANS Amount],[GL TRANS Posted YN],"
      SQLstatement = SQLstatement & " [GL TRANS Description],[GL TRANS Source],[GL TRANS System Generated])"
    
    Dim TempStr As String
    
    ' gl post date
    If PostDate% = 1 Then
      TempStr = Format(Now, "Short Date")
    Else
      TempStr = rsPayment("AP PAY Transaction Date")
    End If
    
      SQLstatement = SQLstatement & " VALUES ('CASH PAY " & CheckNo$ & "-" & rsPayment("AP PAY Bank Account") & "','Cash Payment',#" & TempStr & "#,"
      
    If rsVendor.BOF And rsVendor.EOF Then
      refr$ = "Unknown"
    Else
      refr$ = rsVendor("AP VEN Name")
    End If

      SQLstatement = SQLstatement & "'" & refr$ & "'," & CCur(rsPayment("AP PAY Amount")) & ",1,"
      SQLstatement = SQLstatement & "'CASH PAY " & CheckNo$ & "','CASH PAY " & CheckNo$ & "',True)"
      'Debug.Print SQLstatement
      
      db.Execute SQLstatement
    
    'rsGLTrans("GL TRANS Reference") = refr$
    'rsGLTrans("GL TRANS Amount") = rsPayment("AP PAY Amount")
    'rsGLTrans("GL TRANS Posted YN") = 1
    'desc$ = "CASH PAY " & CheckNo$
    'rsGLTrans("GL TRANS Description") = desc$
    'rsGLTrans("GL TRANS Source") = "CASH PAY " & CheckNo$
    'rsGLTrans("GL TRANS System Generated") = True
  'rsGLTrans.Update
  Dim rsGLTrans As ADODB.Recordset
  Set rsGLTrans = New ADODB.Recordset
  rsGLTrans.Open "SELECT [GL TRANS Number] FROM [GL Transaction]WHERE [GL TRANS Document #]='" & "CASH PAY " & CheckNo$ & "-" & rsPayment("AP PAY Bank Account") & "'", db, adOpenKeyset, adLockReadOnly, adCmdText
      NewNumber& = rsGLTrans("GL TRANS Number")
  rsGLTrans.Close
  Set rsGLTrans = Nothing
  ' write GL Transaction Detail

  'Loop through line items
  Dim rsCross As ADODB.Recordset
  Set rsCross = New ADODB.Recordset
  'rsCross.Open "SELECT * FROM [AP Payment Invoice Cross Reference] WHERE [AP CROSS Payment ID] = " & rsPayment("AP PAY ID"), db, adOpenStatic, adLockOptimistic, adCmdText
  rsCross.Open "SELECT [AP CROSS Applied Amount], [AP CROSS Discount Taken],[AP CROSS Write Off Amount]," & _
  "[AP CROSS Payed ID] FROM [AP Payment Invoice Cross Reference] WHERE " & _
  "[AP CROSS Payment ID] = " & rsPayment("AP PAY ID"), db, adOpenKeyset, adLockOptimistic, adCmdText

  Dim TotalApplied@
  TotalApplied@ = 0

  'On Error Resume Next
  'If Err = 0 Then
  If rsCross.RecordCount > 0 Then
    rsCross.MoveFirst
    Do While Not rsCross.EOF
      ' only process records with payments, discounts or writeoffs
      If rsCross("AP CROSS Applied Amount") > 0 Or rsCross("AP CROSS Discount Taken") > 0 Then

        ' process payments
        If rsCross("AP CROSS Applied Amount") > 0 Then
          
          TotalApplied@ = TotalApplied@ + rsCross("AP CROSS Applied Amount")

'          ' update GL for payment
          '-----------------------------------------------------------------------
          ' Payment GL Affected Accounts
          '
          '                  Debit   Credit   Source
          '                  -----   ------   ------
          ' CASH               X              Bank - Cash Acct
          ' AP                         X      Pref - Sales
          '-----------------------------------------------------------------------

          ' Debits
          ' Cash Receipt
            SQLstatement = "INSERT INTO [GL Work Detail]"
            SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
            SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & rsPayment("AP PAY Bank Account") & "" & "',0," & rsCross("AP CROSS Applied Amount") & ")"
            db.Execute SQLstatement
          
          'rsGLWorkDetail.AddNew
          '  rsGLWorkDetail("GW TRANSD Number") = NewNumber&
          '  rsGLWorkDetail("GW TRANSD Account") = rsPayment("AP PAY Bank Account")
          '  rsGLWorkDetail("GW TRANSD Debit Amount") = 0
          '  rsGLWorkDetail("GW TRANSD Credit Amount") = rsCross("AP CROSS Applied Amount")
          '  rsGLWorkDetail("GW TRANSD Project") = ""
          'rsGLWorkDetail.Update

          ' Credits
          ' AR
            SQLstatement = "INSERT INTO [GL Work Detail]"
            SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
            SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & rsCompany("SYS COM Purchase AP Acct") & "" & "'," & rsCross("AP CROSS Applied Amount") & ",0)"
            db.Execute SQLstatement
          
         'rsGLWorkDetail.AddNew
         '   rsGLWorkDetail("GW TRANSD Number") = NewNumber&
         '   rsGLWorkDetail("GW TRANSD Account") = rsCompany("SYS COM Purchase AP Acct")
         '   rsGLWorkDetail("GW TRANSD Debit Amount") = rsCross("AP CROSS Applied Amount")
         '   rsGLWorkDetail("GW TRANSD Credit Amount") = 0
         '   rsGLWorkDetail("GW TRANSD Project") = ""
         ' rsGLWorkDetail.Update
          ' update GL for payment

        End If ' end process payments
        
        ' process discount amounts
        If rsCross("AP CROSS Discount Taken") > 0 Then

 '         ' update GL for discount
          '-----------------------------------------------------------------------
          ' Discount GL Affected Accounts
          '
          '                  Debit   Credit   Source
          '                  -----   ------   ------
          ' Discount           X              Pref - Sales
          ' AR                         X      Pref - Sales
          '-----------------------------------------------------------------------

          ' Debits
          ' Discount
            SQLstatement = "INSERT INTO [GL Work Detail]"
            SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
            SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & rsCompany("SYS COM Purchase Discount Acct") & "" & "',0," & rsCross("AP CROSS Discount Taken") & ")"
            db.Execute SQLstatement
          
          'rsGLWorkDetail.AddNew
          '  rsGLWorkDetail("GW TRANSD Number") = NewNumber&
          '  rsGLWorkDetail("GW TRANSD Account") = rsCompany("SYS COM Purchase Discount Acct")
          '  rsGLWorkDetail("GW TRANSD Debit Amount") = 0
          '  rsGLWorkDetail("GW TRANSD Credit Amount") = rsCross("AP CROSS Discount Taken")
          '  rsGLWorkDetail("GW TRANSD Project") = ""
          'rsGLWorkDetail.Update

          ' Credits
          ' AP
            SQLstatement = "INSERT INTO [GL Work Detail]"
            SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
            SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & rsCompany("SYS COM Purchase AP Acct") & "" & "'," & rsCross("AP CROSS Discount Taken") & ",0)"
            db.Execute SQLstatement
          
          'rsGLWorkDetail.AddNew
          '  rsGLWorkDetail("GW TRANSD Number") = NewNumber&
          '  rsGLWorkDetail("GW TRANSD Account") = rsCompany("SYS COM Purchase AP Acct")
          '  rsGLWorkDetail("GW TRANSD Debit Amount") = rsCross("AP CROSS Discount Taken")
          '  rsGLWorkDetail("GW TRANSD Credit Amount") = 0
          '  rsGLWorkDetail("GW TRANSD Project") = ""
          'rsGLWorkDetail.Update
          ' update GL for discount

        End If ' end process discount amounts
        
      End If
      rsCross.MoveNext
    Loop
  End If

  ' handle Unapplied Payments or Payments On Account
  If TotalApplied@ < rsPayment("AP PAY Amount") Then

    ' Credits
    ' Cash Receipt
      SQLstatement = "INSERT INTO [GL Work Detail]"
      SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
      SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & rsPayment("AP PAY Bank Account") & "" & "',0," & rsPayment("AP PAY Amount") - TotalApplied@ & ")"
      db.Execute SQLstatement
          
    'rsGLWorkDetail.AddNew
    '  rsGLWorkDetail("GW TRANSD Number") = NewNumber&
    '  rsGLWorkDetail("GW TRANSD Account") = rsPayment("AP PAY Bank Account")
    '  rsGLWorkDetail("GW TRANSD Debit Amount") = 0
    '  rsGLWorkDetail("GW TRANSD Credit Amount") = rsPayment("AP PAY Amount") - TotalApplied@
    '  rsGLWorkDetail("GW TRANSD Project") = ""
    'rsGLWorkDetail.Update

    ' Debits
    ' AP
      SQLstatement = "INSERT INTO [GL Work Detail]"
      SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
      SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & rsCompany("SYS COM Purchase AP Acct") & "" & "'," & rsPayment("AP PAY Amount") - TotalApplied@ & ",0)"
      db.Execute SQLstatement
          
    'rsGLWorkDetail.AddNew
    '  rsGLWorkDetail("GW TRANSD Number") = NewNumber&
    '  rsGLWorkDetail("GW TRANSD Account") = rsCompany("SYS COM Purchase AP Acct")
    '  rsGLWorkDetail("GW TRANSD Debit Amount") = rsPayment("AP PAY Amount") - TotalApplied@
    '  rsGLWorkDetail("GW TRANSD Credit Amount") = 0
    '  rsGLWorkDetail("GW TRANSD Project") = ""
    'rsGLWorkDetail.Update
    ' update GL for payment

  End If
  ' end of handle Unapplied Payments or Payments On Account

  ' update vendor stats
    rsVendor("AP VEN Payments YTD") = rsVendor("AP VEN Payments YTD") + rsPayment("AP PAY Amount")
    rsVendor("AP VEN Payments Lifetime") = rsVendor("AP VEN Payments Lifetime") + rsPayment("AP PAY Amount")
    ' Update current Balance - if not Paid in full
    CurrentBalance@ = IIf(IsNull(rsVendor("AP VEN Financial Period 1")), 0, rsVendor("AP VEN Financial Period 1"))
    CurrentBalance@ = CurrentBalance@ - rsPayment("AP PAY Amount")
    rsVendor("AP VEN Financial Period 1") = CurrentBalance@
  
  rsVendor.Update
  rsVendor.Close
  Set rsVendor = Nothing

  ' post GL entry
  Dim Success%
  Success% = PostGLWorkDetail(TranDate, NewNumber&, db)
  If Success% = False Then
    PostPayments = False
  Else
    PostPayments = True
  End If

Postpayments_Exit:
  'rsVendor.Close
  'Set rsVendor = Nothing
  rsCross.Close
  Set rsCross = Nothing
  'rsGLTrans.Close
  'Set rsGLTrans = Nothing
  rsCompany.Close
  Set rsCompany = Nothing
  'rsGLWorkDetail.Close
  'Set rsGLWorkDetail = Nothing
  rsPayment.Close
  Set rsPayment = Nothing
If Currentdb = True Then
    db.Close
    Set db = Nothing
End If
  Exit Function

PostPayments_Error:
  Call ErrorLog("Purchase Module", "PostPayments", Now, Err.Number, Err.Description, True, db)
  PostPayments% = False
  Exit Function

AllreadyPosted:
  rsPayment.Close
  Set rsPayment = Nothing
If Currentdb = True Then
    db.Close
    Set db = Nothing
End If
  Exit Function

AllreadyClosed:
  rsCompany.Close
  Set rsCompany = Nothing
  'rsGLWorkDetail.Close
  'Set rsGLWorkDetail = Nothing
  rsPayment.Close
  Set rsPayment = Nothing
If Currentdb = True Then
  db.Close
  Set db = Nothing
End If
  Exit Function

End Function
Function PostPOCreditMemo(DocumentKey&, intShowError As Integer, db As ADODB.Connection)

  'On Error GoTo PostPOCreditMemo_Error

  'Dim db As ADODB.Connection
  'Set db = New ADODB.Connection
  'db.CursorLocation = adUseServer
  'db.Open gblADOProvider
  
  Dim msg$
  Dim title$

  Dim rsCompany As ADODB.Recordset
  Set rsCompany = New ADODB.Recordset
  rsCompany.Open "SELECT  [SYS COM Purchase AP Acct],[SYS COM Purchase Discount Acct]," & _
  "[SYS COM GL Post By Date],[SYS COM Purchase Inventory Acct],[SYS COM Purchase Freight Acct]," & _
  "[SYS COM Purchase Misc Acct] FROM [SYS Company]", db, adOpenKeyset, adLockOptimistic, admcdtext
  rsCompany.MoveFirst

  ' clear any GL Work records
  'Dim cmdtemp As ADODB.Recordset
  'Set cmdtemp = New ADODB.Recordset
  'cmdtemp.Open "DELETE DISTINCTROW * FROM [GL Work Detail]", db, , , adCmdText
'  cmdtemp.Close
  'Set cmdtemp = Nothing
  db.Execute "DELETE DISTINCTROW * FROM [GL Work Detail]"
  
  'Dim rsGLWorkDetail As ADODB.Recordset
  'Set rsGLWorkDetail = New ADODB.Recordset
  'rsGLWorkDetail.Open "[GL Work Detail]", db, adOpenStatic, adLockOptimistic, adCmdTable
  
  Dim rsPurchase As ADODB.Recordset
  Set rsPurchase = New ADODB.Recordset
  rsPurchase.Open "SELECT [AP PO Document Type],[AP PO Date],[AP PO Vendor ID],[AP PO Amount Paid]," & _
  "[AP PO Payment Method],[AP PO Check Number],[AP PO Check Acct ID],[AP PO Document No]," & _
  "[AP PO Ext Document No],[AP PO vendor Name],[AP PO Description],[AP PO Total Amount]," & _
  "[AP PO Shipping],[AP PO Misc Charges],[AP PO Vendor Invoice No],[AP PO Discount Amt] FROM [AP Purchase] " & _
  "WHERE [AP PO Document No]=" & DocumentKey&, db, adOpenKeyset, adLockOptimistic, adCmdText

  ' first lets get the Credit Memo
  'rsPurchase.Index = "PrimaryKey"
  'rsPurchase.Seek DocumentKey&
  'rsPurchase.MoveFirst
  'rsPurchase.Find "[AP PO Document No]='" & DocumentKey& & "'"

  'Post by 1 - system date or 2 - Transaction date?
  Dim PostDate%
  PostDate% = rsCompany("SYS COM GL Post By Date")

  'Set Purchase Type
  Dim PurchaseType$
  PurchaseType$ = rsPurchase("AP PO Document Type")

  'Set Post Date
  Dim TranDate As Variant
  If PostDate% = 1 Then
    TranDate = DateValue(Format(Now, "Short Date"))
  Else
    TranDate = DateValue(rsPurchase("AP PO Date"))
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
    GoTo UnableToPostPOCreditMemoHere
  End If

  'On Error GoTo PostPOCreditMemo_Error

  ' update vendor stats
  Dim rsVendor As ADODB.Recordset
  Set rsVendor = New ADODB.Recordset
  rsVendor.Open "SELECT [AP VEN ID], [AP VEN Name], [AP VEN Payments YTD]," & _
  "[AP VEN Payment Number Lifetime],[AP VEN Payment Number YTD],[AP VEN Purchase YTD], " & _
  "[AP VEN Purchase Lifetime],[AP VEN Purchase Number Lifetime],[AP VEN Purchase Number YTD]," & _
  "[AP VEN Payments Lifetime], [AP VEN Financial Period 1] FROM [AP Vendor] " & _
  "WHERE [AP VEN ID]='" & rsPurchase("AP PO Vendor ID") & "'", db, adOpenKeyset, adLockOptimistic, adCmdText

  Dim CurrentBalance@

    rsVendor("AP VEN Purchase YTD") = rsVendor("AP VEN Purchase YTD") - rsPurchase("AP PO Total Amount")
    rsVendor("AP VEN Purchase Lifetime") = rsVendor("AP VEN Purchase Lifetime") - rsPurchase("AP PO Total Amount")
    rsVendor("AP VEN Purchase Number Lifetime") = rsVendor("AP VEN Purchase Number Lifetime") - 1
    rsVendor("AP VEN Purchase Number YTD") = rsVendor("AP VEN Purchase Number YTD") - 1
    CurrentBalance@ = IIf(IsNull(rsVendor("AP VEN Financial Period 1")), 0, rsVendor("AP VEN Financial Period 1"))
    CurrentBalance@ = CurrentBalance@ - rsPurchase("AP PO Total Amount")
    rsVendor("AP VEN Financial Period 1") = CurrentBalance@
  rsVendor.Update
  rsVendor.Close
  Set rsVendor = Nothing
  
  '--------------------------------------------------
  ' New AP Cross Payment and AP Payment Header

  'Dim rsAPPaymentHeader As ADODB.Recordset
  'Set rsAPPaymentHeader = New ADODB.Recordset
  'rsAPPaymentHeader.Open "[AP Payment Header]", db, adOpenStatic, adLockOptimistic, adCmdTable
  
  'Dim rsAPCross As ADODB.Recordset
  'Set rsAPCross = New ADODB.Recordset
  'rsAPCross.Open "[AP Payment Invoice Cross Reference]", db, adOpenStatic, adLockOptimistic, adCmdTable

  'Dim PaymentID&

    SQLstatement = "INSERT INTO [AP Payment Header]"
    SQLstatement = SQLstatement & " ([AP PAY Type],[AP PAY Check No],[AP PAY Vendor No]," & _
    "[AP PAY Transaction Date],[AP PAY Amount],[AP PAY UnApplied Amount],[AP PAY Credit YN]," & _
    "[AP PAY Bank Account],[AP PAY Status],[AP PAY Class],[AP PAY Posted YN],[AP PAY Void],[AP PAY Cleared])"
    
    '  write Payment Header
    SQLstatement = SQLstatement & " VALUES ('Credit Memo','" & _
    "CM " & rsPurchase("AP PO Vendor Invoice No") & "','" & rsPurchase("AP PO Vendor ID") & "',#" & _
    rsPurchase("AP PO Date") & "#," & rsPurchase("AP PO Total Amount") & ","
    SQLstatement = SQLstatement & rsPurchase("AP PO Total Amount") & ",True,'None',"
    SQLstatement = SQLstatement & "'Posted',0,True,False,False)"
  
  'rsAPPaymentHeader.AddNew
  '  rsAPPaymentHeader("AP PAY Type") = "Credit Memo"
  '  rsAPPaymentHeader("AP PAY Check No") = "CM " & rsPurchase("AP PO Ext Document No")
  '  rsAPPaymentHeader("AP PAY Vendor No") = rsPurchase("AP PO Vendor ID")
  '  rsAPPaymentHeader("AP PAY Transaction Date") = rsPurchase("AP PO Date")
  '  rsAPPaymentHeader("AP PAY Amount") = rsPurchase("AP PO Total Amount")
  '  rsAPPaymentHeader("AP PAY UnApplied Amount") = rsPurchase("AP PO Total Amount")
  '  rsAPPaymentHeader("AP PAY Credit YN") = True
  '  rsAPPaymentHeader("AP PAY Bank Account") = "None"
  '  rsAPPaymentHeader("AP PAY Status") = "Posted"
  '  rsAPPaymentHeader("AP PAY Void") = False  'Changed NSF to Void
  '  rsAPPaymentHeader("AP PAY Class") = 0 '"CreditMemo"
  '  rsAPPaymentHeader("AP PAY Cleared") = False
  '  rsAPPaymentHeader("AP PAY Posted YN") = True
  'rsAPPaymentHeader.Update
      'PaymentID& = rsARPaymentHeader("AR PAY ID")
      Dim rsAPPaymentHeader As ADODB.Recordset
      Set rsAPPaymentHeader = New ADODB.Recordset
      rsAPPaymentHeader.Open "SELECT [AP PAY ID] FROM [AP Payment Header] WHERE [AP PAY Check No]='CM " & rsPurchase("AP PO Vendor Invoice No") & "'", db, adOpenKeyset, adLockReadOnly, adCmdText
      If rsAPPaymentHeader.RecordCount > 1 Then
        rsAPPaymentHeader.MoveLast
        PaymentID& = rsAPPaymentHeader("AP PAY ID")
      Else
        PaymentID& = rsAPPaymentHeader("AP PAY ID")
      End If
      rsAPPaymentHeader.Close
      Set rsAPPaymentHeader = Nothing
  ' end of write payment header

  ' End of New AP Cross Payment and AP Payment Header
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
      SQLstatement = SQLstatement & " [GL TRANS Description],[GL TRANS Source],[GL TRANS System Generated])"
    
    'rsGLTrans("GL TRANS Document #") = "POCM " & rsPurchase("AP PO Ext Document No")
    
    ' gl post date
    Dim TempStr As String
    If PostDate% = 1 Then
      TempStr = Format(Now, "Short Date")
    Else
      TempStr = rsPurchase("AP PO Date")
    End If
    
    'rsGLTrans("GL TRANS Type") = "Credit Memo"
    SQLstatement = SQLstatement & " VALUES ('POCM " & rsPurchase("AP PO Ext Document No") & "','Credit Memo',#" & TempStr & "#,"

    refr$ = rsPurchase("AP PO vendor Name")
    desc$ = IIf(IsNull(rsPurchase("AP PO Description")), "", rsPurchase("AP PO Description"))
    If Len(Trim$(desc$)) = 0 Then
      desc$ = "POCM " & rsPurchase("AP PO Ext Document No")
    End If
    
      SQLstatement = SQLstatement & "'" & refr$ & "'," & rsPurchase("AP PO Total Amount") & ",1,"
      SQLstatement = SQLstatement & "'" & desc$ & "','POCM " & rsPurchase("AP PO Ext Document No") & "',True)"
      'Debug.Print SQLstatement
      
      db.Execute SQLstatement
    
    'rsGLTrans("GL TRANS Reference") = refr$
    'rsGLTrans("GL TRANS Amount") = rsPurchase("AP PO Total Amount")
    'rsGLTrans("GL TRANS Posted YN") = 1
    'rsGLTrans("GL TRANS Description") = desc$
    'rsGLTrans("GL TRANS Source") = "POCM " & rsPurchase("AP PO Ext Document No")
    'rsGLTrans("GL TRANS System Generated") = True
  'rsGLTrans.Update
  Dim rsGLTrans As ADODB.Recordset
  Set rsGLTrans = New ADODB.Recordset
  rsGLTrans.Open "SELECT [GL TRANS Number] FROM [GL Transaction] WHERE [GL TRANS Document #]='" & "POCM " & rsPurchase("AP PO Ext Document No") & "'", db, adOpenKeyset, adLockReadOnly, adCmdText
      NewNumber& = rsGLTrans("GL TRANS Number")
  rsGLTrans.Close
  Set rsGLTrans = Nothing

  ' update GL
  '-----------------------------------------------------------------------
  ' Voucher without Payment GL Affected Accounts
  '
  '                  Debit   Credit   Source
  '                  -----   ------   ------
  ' Voucher Item               X       AP Detail
  ' AP                 X              Pref - Purchases
  ' Discount Amt       X              Pref - Purchase Discount Acct xxx 10/23/95
  ' Misc Charges               X      Pref - Misc Charges
  ' Freight                    X      Pref - Freight
  '
  '
  ' Notes:
  ' Each Detail Item is processed and the GL Acct is Retrieved.
  '-----------------------------------------------------------------------
  
  'Debits
  ' AP
  'rsGLWorkDetail.AddNew
      SQLstatement = "INSERT INTO [GL Work Detail]"
      SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
      SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & rsCompany("SYS COM Purchase AP Acct") & "'," & rsPurchase("AP PO Total Amount") & ",0)"
      db.Execute SQLstatement
    
  '  rsGLWorkDetail("GW TRANSD Number") = NewNumber&
  '  rsGLWorkDetail("GW TRANSD Account") = rsCompany("SYS COM Purchase AP Acct")
  '  rsGLWorkDetail("GW TRANSD Debit Amount") = rsPurchase("AP PO Total Amount")
  '  rsGLWorkDetail("GW TRANSD Credit Amount") = 0
  '  rsGLWorkDetail("GW TRANSD Project") = ""
  'rsGLWorkDetail.Update
  
  'Discount Amount
  If rsPurchase("AP PO Discount Amt") > 0 Then
      SQLstatement = "INSERT INTO [GL Work Detail]"
      SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
      SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & rsCompany("SYS COM Purchase Discount Acct") & "'," & rsPurchase("AP PO Discount Amt") & ",0)"
      db.Execute SQLstatement
    'rsGLWorkDetail.AddNew
    '  rsGLWorkDetail("GW TRANSD Number") = NewNumber&
    '  rsGLWorkDetail("GW TRANSD Account") = rsCompany("SYS COM Purchase Discount Acct")
    '  rsGLWorkDetail("GW TRANSD Debit Amount") = rsPurchase("AP PO Discount Amt")
    '  rsGLWorkDetail("GW TRANSD Credit Amount") = 0
    '  rsGLWorkDetail("GW TRANSD Project") = ""
    'rsGLWorkDetail.Update
  End If

  ' Credit
  
  ' Freight Expense
  If rsPurchase("AP PO Shipping") > 0 Then
      SQLstatement = "INSERT INTO [GL Work Detail]"
      SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
      SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & rsCompany("SYS COM Purchase Freight Acct") & "',0," & rsPurchase("AP PO Shipping") & ")"
      db.Execute SQLstatement
    'rsGLWorkDetail.AddNew
    '  rsGLWorkDetail("GW TRANSD Number") = NewNumber&
    '  rsGLWorkDetail("GW TRANSD Account") = rsCompany("SYS COM Purchase Freight Acct")
    '  rsGLWorkDetail("GW TRANSD Debit Amount") = 0
    '  rsGLWorkDetail("GW TRANSD Credit Amount") = rsPurchase("AP PO Shipping")
    '  rsGLWorkDetail("GW TRANSD Project") = ""
    'rsGLWorkDetail.Update
  End If

  ' Misc Charges
  If rsPurchase("AP PO Misc Charges") > 0 Then
      SQLstatement = "INSERT INTO [GL Work Detail]"
      SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
      SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & rsCompany("SYS COM Purchase Misc Acct") & "',0," & rsPurchase("AP PO Misc Charges") & ")"
      db.Execute SQLstatement
    'rsGLWorkDetail.AddNew
    '  rsGLWorkDetail("GW TRANSD Number") = NewNumber&
    '  rsGLWorkDetail("GW TRANSD Account") = rsCompany("SYS COM Purchase Misc Acct")
    '  rsGLWorkDetail("GW TRANSD Debit Amount") = 0
    '  rsGLWorkDetail("GW TRANSD Credit Amount") = rsPurchase("AP PO Misc Charges")
    '  rsGLWorkDetail("GW TRANSD Project") = ""
    'rsGLWorkDetail.Update
  End If
  
  ' Detail Item Increase
  Dim Longer&
  Longer& = 0
  Dim rsDetail As ADODB.Recordset
  Dim APDetailAcct$
  Set rsDetail = New ADODB.Recordset
  rsDetail.Open "SELECT [AP POD Posting Account],[AP POD Item ID],[AP POD Item Total] " & _
  "FROM [AP Purchase Detail] where [AP POD Document No] = " & rsPurchase("AP PO Document No"), db, adOpenKeyset, adLockOptimistic, adCmdText

  If rsDetail.RecordCount = 0 Then
    ' no detail data
  Else
    rsDetail.MoveFirst
    ' may have detail data
    Do While rsDetail.EOF = False
      ' use Account in detail 1st
      APDetailAcct$ = IIf(IsNull(rsDetail("AP POD Posting Account")), "", rsDetail("AP POD Posting Account"))

      ' use Account in Vendor 2nd
      If Len(APDetailAcct$) = 0 Then
        Dim VendorKey$
        VendorKey$ = IIf(IsNull(rsPurchase("AP PO Vendor ID")), "", rsPurchase("AP PO Vendor ID"))
        If Len(VendorKey$) > 0 Then
          'rsVendor.MoveFirst
          'rsVendor.Index = "PrimaryKey"
          'rsVendor.Seek VendorKey$ AP VEN ID
          'rsVendor.MoveFirst
          'rsVendor.Find "[AP VEN ID]='" & VendorKey$ & "'"
          'If rsVendor.BOF And rsVendor.EOF Then
          Set rsVendor = New ADODB.Recordset
          rsVendor.Open "SELECT [AP VEN Default GL] FROM [AP Vendor] WHERE [AP VEN ID]='" & VendorKey$ & "'", db, adOpenKeyset, adLockReadOnly, adcmdtex
            If rsVendor.RecordCount = 0 Then
            Else
              APDetailAcct$ = IIf(IsNull(rsVendor("AP VEN Default GL")), "", rsVendor("AP VEN Default GL"))
            End If
          rsVendor.Close
          Set rsVendor = Nothing
        End If
      End If

      'Use Misc Purchase Acct 3rd
      If Len(APDetailAcct$) = 0 Then APDetailAcct$ = rsCompany("SYS COM Purchase Misc Acct")
      SQLstatement = "INSERT INTO [GL Work Detail]"
      SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
      SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & APDetailAcct$ & "',0," & rsDetail("AP POD Item Total") & ")"
      db.Execute SQLstatement
      
      'rsGLWorkDetail.AddNew
      '  rsGLWorkDetail("GW TRANSD Number") = NewNumber&
      '  rsGLWorkDetail("GW TRANSD Account") = APDetailAcct$
      '  rsGLWorkDetail("GW TRANSD Debit Amount") = 0
      '  rsGLWorkDetail("GW TRANSD Credit Amount") = rsDetail("AP POD Item Total")
      '  rsGLWorkDetail("GW TRANSD Project") = ""
      'rsGLWorkDetail.Update

      rsDetail.MoveNext
    Loop
  rsDetail.Close
  Set rsDetail = Nothing
  End If
  
  ' post GL entry for this Credit Memo
  Dim Success%
  Success% = PostGLWorkDetail(TranDate, NewNumber&, db)
  If Success% = False Then
    MsgBox "An error occurred writing GL Transaction!", , "Error"
    PostPOCreditMemo = False
    Exit Function
  End If

  PostPOCreditMemo = True

PostPOCreditMemo_Exit:
   'On Error Resume Next
   rsCompany.Close
   Set rsCompany = Nothing
   rsPurchase.Close
   Set rsPurchase = Nothing
   'rsGLWorkDetail.Close
   'Set rsGLWorkDetail = Nothing
   'rsGLTrans.Close
   'Set rsGLTrans = Nothing
   'rsVendor.Close
   'Set rsVendor = Nothing
   'rsAPCross.Close
   'Set rsAPCross = Nothing
   'rsAPPaymentHeader.Close
   'Set rsAPPaymentHeader = Nothing
   'rsDetail.Close
   'Set rsDetail = Nothing
   'db.Close
   'Set db = Nothing
  Exit Function

PostPOCreditMemo_Error:
  Call ErrorLog("Purchase Module", "PostPOCreditMemo", Now, Err.Number, Err.Description, intShowError, db)
  PostPOCreditMemo = False
   'On Error Resume Next
   rsCompany.Close
   Set rsCompany = Nothing
   rsPurchase.Close
   Set rsPurchase = Nothing
   'rsGLWorkDetail.Close
   'Set rsGLWorkDetail = Nothing
   'rsGLTrans.Close
   'Set rsGLTrans = Nothing
   'rsVendor.Close
   'Set rsVendor = Nothing
   'rsAPCross.Close
   'Set rsAPCross = Nothing
   'rsAPPaymentHeader.Close
   'Set rsAPPaymentHeader = Nothing
   'rsDetail.Close
   'Set rsDetail = Nothing
   'db.Close
   'Set db = Nothing
  Exit Function

UnableToPostPOCreditMemoHere:
  PostPOCreditMemo = False
  'On Error Resume Next
   rsCompany.Close
   Set rsCompany = Nothing
   rsPurchase.Close
   Set rsPurchase = Nothing
   'rsGLWorkDetail.Close
   'Set rsGLWorkDetail = Nothing
   'rsGLTrans.Close
   'Set rsGLTrans = Nothing
   'rsVendor.Close
   'Set rsVendor = Nothing
   'rsAPCross.Close
   'Set rsAPCross = Nothing
   'rsAPPaymentHeader.Close
   'Set rsAPPaymentHeader = Nothing
   'rsDetail.Close
   'Set rsDetail = Nothing
   'db.Close
   'Set db = Nothing
  Exit Function

End Function

Function PostReceiving(DocumentKey&, intShowError As Integer, db As ADODB.Connection)

  'On Error GoTo PostReceiving_Error

  'Dim db As ADODB.Connection
  'Set db = New ADODB.Connection
  'db.CursorLocation = adUseServer
  'db.Open gblADOProvider

  Dim msg$
  Dim title$

  Dim rsCompany As ADODB.Recordset
  Set rsCompany = New ADODB.Recordset
  'rsCompany.Open "[SYS Company]", db, adOpenStatic, adLockOptimistic, adCmdTable
  rsCompany.Open "SELECT  [SYS COM Purchase AP Acct],[SYS COM Purchase Discount Acct]," & _
  "[SYS COM GL Post By Date],[SYS COM Purchase Inventory Acct],[SYS COM Purchase Freight Acct]," & _
  "[SYS COM Purchase Misc Acct] FROM [SYS Company]", db, adOpenKeyset, adLockOptimistic, admcdtext
  rsCompany.MoveFirst

  ' clear any GL Work records
  'Dim cmdtemp As ADODB.Recordset
  'cmdtemp.Open "DELETE DISTINCTROW * FROM [GL Work Detail]", db, , , adCmdText
  db.Execute "DELETE DISTINCTROW * FROM [GL Work Detail]"
  'cmdtemp.Close
  'Set cmdtemp = Nothing

  'Dim rsGLWorkDetail As ADODB.Recordset
  'Set rsGLWorkDetail = New ADODB.Recordset
  'rsGLWorkDetail.Open "[GL Work Detail]", db, adOpenStatic, adLockOptimistic, adCmdTable

  Dim rsPurchase As ADODB.Recordset
  Set rsPurchase = New ADODB.Recordset
  rsPurchase.Open "SELECT [AP PO Document Type],[AP PO Date],[AP PO Vendor ID],[AP PO Amount Paid]," & _
  "[AP PO Payment Method],[AP PO Check Number],[AP PO Check Acct ID],[AP PO Document No]," & _
  "[AP PO Ext Document No],[AP PO vendor Name],[AP PO Description],[AP PO Total Amount]," & _
  "[AP PO Shipping],[AP PO Misc Charges],[AP PO Discount Amt] FROM [AP Purchase] " & _
  "WHERE [AP PO Document No]=" & DocumentKey&, db, adOpenKeyset, adLockOptimistic, adCmdText
  
  'rsPurchase.Index = "PrimaryKey"
  'rsPurchase.MoveFirst
  'rsPurchase.Find "[AP PO Document No]=" & DocumentKey&

  'Post by 1 - system date or 2 - Transaction date?
  Dim PostDate%
  PostDate% = rsCompany("SYS COM GL Post By Date")

  'Set Purchase Type
  Dim PurchaseType$
  PurchaseType$ = rsPurchase("AP PO Document Type")

  'Set Post Date
  Dim TranDate As Variant
  If PostDate% = 1 Then
    TranDate = DateValue(Format(Now, "Short Date"))
  Else
    TranDate = DateValue(rsPurchase("AP PO Date"))
  End If

  'Verify period can be posted to
  'Send TranDate
  'Return PeriodToPost and PeriodClosed
  Dim PeriodToPost%
  Dim PeriodClosed%
  '<<<----took almost 6 second to execute the above code ---- do something (SQL statement)
  Call VerifyPeriod(TranDate, PeriodToPost%, PeriodClosed%, db)
  If PeriodToPost% = 0 Then
    PostReceiving = False
    MsgBox "The transaction date " & FormatDate(CDate(TranDate)) & " is not within the Accounting Periods", vbCritical, "Error"
  Exit Function
  End If
  
  'Is period open?
  If PeriodClosed% = True Then '<<<--- 11 seconds
    MsgBox "Unable to post transaction to a closed period.", , "Post Invoice Error"
    GoTo UnableToPostPOHere
  End If

  'On Error GoTo PostReceiving_Error

  ' update vendor stats
  Dim rsVendor As ADODB.Recordset
  Set rsVendor = New ADODB.Recordset
  'rsVendor.Open "SELECT * FROM [AP Vendor] where [AP VEN ID] = '" & rsPurchase("AP PO Vendor ID") & "'", db, adOpenStatic, adLockOptimistic, adCmdText
  rsVendor.Open "SELECT [AP VEN ID], [AP VEN Name], [AP VEN Payments YTD]," & _
  "[AP VEN Payment Number Lifetime],[AP VEN Payment Number YTD],[AP VEN Purchase YTD], " & _
  "[AP VEN Purchase Lifetime],[AP VEN Purchase Number Lifetime],[AP VEN Purchase Number YTD]," & _
  "[AP VEN Payments Lifetime], [AP VEN Financial Period 1] FROM [AP Vendor] " & _
  "WHERE [AP VEN ID]='" & rsPurchase![AP PO Vendor ID] & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
  'rsVendor.Index = "PrimaryKey"
  'rsVendor.Seek VendorID$

  Dim CurrentBalance@

    If rsPurchase("AP PO Amount Paid") > 0 Then
      rsVendor("AP VEN Payments YTD") = rsVendor("AP VEN Payments YTD") + rsPurchase("AP PO Amount Paid")
      rsVendor("AP VEN Payments Lifetime") = rsVendor("AP VEN Payments Lifetime") + rsPurchase("AP PO Amount Paid")
      rsVendor("AP VEN Payment Number Lifetime") = rsVendor("AP VEN Payment Number Lifetime") + 1
      rsVendor("AP VEN Payment Number YTD") = rsVendor("AP VEN Payment Number YTD") + 1
    End If
    rsVendor("AP VEN Purchase YTD") = rsVendor("AP VEN Purchase YTD") + rsPurchase("AP PO Total Amount")
    rsVendor("AP VEN Purchase Lifetime") = rsVendor("AP VEN Purchase Lifetime") + rsPurchase("AP PO Total Amount")
    rsVendor("AP VEN Purchase Number Lifetime") = rsVendor("AP VEN Purchase Number Lifetime") + 1
    rsVendor("AP VEN Purchase Number YTD") = rsVendor("AP VEN Purchase Number YTD") + 1
    ' Update current Balance - if not Paid in full
    If rsPurchase("AP PO Total Amount") - rsPurchase("AP PO Amount Paid") > 0 Then
      CurrentBalance@ = IIf(IsNull(rsVendor("AP VEN Financial Period 1")), 0, rsVendor("AP VEN Financial Period 1"))
      CurrentBalance@ = CurrentBalance@ + (rsPurchase("AP PO Total Amount") - rsPurchase("AP PO Amount Paid"))
      rsVendor("AP VEN Financial Period 1") = CurrentBalance@
    End If
  rsVendor.Update
  rsVendor.Close
  Set rsVendor = Nothing

  Dim PaymentID&

  If rsPurchase("AP PO Amount Paid") > 0 Then
    'Dim rsAPPaymentHeader As ADODB.Recordset
    'Set rsAPPaymentHeader = New ADODB.Recordset
    'rsAPPaymentHeader.Open "[AP Payment Header]", db, adOpenStatic, adLockOptimistic, adCmdTable
    
    'Dim rsAPCross As ADODB.Recordset
    'Set rsAPCross = New ADODB.Recordset
    'rsAPCross.Open "[AP Payment Invoice Cross Reference]", db, adOpenStatic, adLockOptimistic, adCmdTable
    
    SQLstatement = "INSERT INTO [AP Payment Header]"
    SQLstatement = SQLstatement & " ([AP PAY Type],[AP PAY Check No],[AP PAY Vendor No]," & _
    "[AP PAY Transaction Date],[AP PAY Amount],[AP PAY UnApplied Amount]" & _
    "[AP PAY Bank Account],[AP PAY Status],[AP PAY Posted YN],[AP PAY Void],[AP PAY Cleared])"
    
    '  write Payment Header
    Dim PayType As String
    'rsAPPaymentHeader.AddNew
      If rsPurchase("AP PO Payment Method") = "Cash" Or rsPurchase("AP PO Payment Method") = "Company Check" Then
        PayType = "Payment"
      Else
        'rsAPPaymentHeader("AP PAY Type") = "Charge"
        PayType = "Charge"
      End If
      
    SQLstatement = SQLstatement & " VALUES ('" & PayType & "','" & _
    rsPurchase("AP PO Check Number") & "','" & rsPurchase("AP PO Vendor ID") & "',#" & _
    rsPurchase("AP PO Date") & "#," & rsPurchase("AP PO Amount Paid") & ",0,'" & _
    rsPurchase("AP PO Check Acct ID") & "','Posted',True,False,False)"
    db.Execute SQLstatement
      
    '  rsAPPaymentHeader("AP PAY Check No") = rsPurchase("AP PO Check Number") & ""
    '  rsAPPaymentHeader("AP PAY Vendor No") = rsPurchase("AP PO Vendor ID") & ""
    '  rsAPPaymentHeader("AP PAY Transaction Date") = rsPurchase("AP PO Date")
    '  rsAPPaymentHeader("AP PAY Amount") = rsPurchase("AP PO Amount Paid")
    '  rsAPPaymentHeader("AP PAY UnApplied Amount") = 0
    '  rsAPPaymentHeader("AP PAY Bank Account") = rsPurchase("AP PO Check Acct ID")
    '  rsAPPaymentHeader("AP PAY Status") = "Posted"
    '  rsAPPaymentHeader("AP PAY Posted YN") = True
    '  rsAPPaymentHeader("AP PAY Void") = False
    '  rsAPPaymentHeader("AP PAY Cleared") = False
      'On Error GoTo PostReceiving_Error
    'rsAPPaymentHeader.Update
      Dim rsAPPaymentHeader As ADODB.Recordset
      Set rsAPPaymentHeader = New ADODB.Recordset
      rsAPPaymentHeader.Open "SELECT [AP PAY ID] FROM [AP Payment Header] WHERE [AP PAY Check No]='" & rsPurchase("AP PO Check Number") & "'", db, adOpenKeyset, adLockReadOnly, adCmdText
      If rsAPPaymentHeader.RecordCount > 1 Then
        rsAPPaymentHeader.MoveLast
        PaymentID& = rsAPPaymentHeader("AP PAY ID")
      Else
        PaymentID& = rsAPPaymentHeader("AP PAY ID")
      End If
      rsAPPaymentHeader.Close
      Set rsAPPaymentHeader = Nothing

    ' end of write payment header
    Dim AmountPaid As Currency
    ' write AR Cross reference Record
      SQLstatement = "INSERT INTO [AP Payment Invoice Cross Reference]"
      SQLstatement = SQLstatement & " ([AP CROSS Payment ID],[AP CROSS Payed ID],"
      SQLstatement = SQLstatement & " [AP CROSS Discount Taken],[AP CROSS Write Off Amount]"
      SQLstatement = SQLstatement & " [AP CROSS Applied Amount],[AP CROSS Cleared])"
      SQLstatement = SQLstatement & " VALUES (" & PaymentID& & ",'" & rsPurchase("AP PO Document No") & "',0,0,"
      
      AmountPaid = IIf(IsNull(rsPurchase("AP PO Amount Paid")), 0, rsPurchase("AP PO Amount Paid"))
      SQLstatement = SQLstatement & AmountPaid & ",False)"
      db.Execute SQLstatement

    'rsAPCross.AddNew
    '  rsAPCross("AP CROSS Payment ID") = PaymentID&
    '  rsAPCross("AP CROSS Payed ID") = rsPurchase("AP PO Document No")
    '  rsAPCross("AP CROSS Discount Taken") = 0
    '  rsAPCross("AP CROSS Write Off Amount") = 0
    '  rsAPCross("AP CROSS Applied Amount") = IIf(IsNull(rsPurchase("AP PO Amount Paid")), 0, rsPurchase("AP PO Amount Paid"))
    '  rsAPCross("AP CROSS Cleared") = False
    'rsAPCross.Update
    ' end of write AR Cross reference Record
  
    ' End of New AP Cross Payment and AP Payment Header
    '--------------------------------------------------
  End If

  ' Inventory Updates
  Dim rsInventory As ADODB.Recordset
  Dim rsDetail As ADODB.Recordset
  
  Set rsInventory = New ADODB.Recordset
  rsInventory.Open "SELECT [INV ITEM Id],[INV ITEM Last Cost],[INV ITEM Average Cost]," & _
  "[INV ITEM Qty On Hand],[INV ITEM Last Cost] FROM [INV Items]", db, adOpenKeyset, adLockOptimistic, adCmdText
  
  Set rsDetail = New ADODB.Recordset
  rsDetail.Open "SELECT [AP POD Item ID],[AP POD Units],[AP POD Unit Cost]," & _
  "[AP POD Total Qty Received] FROM [AP Purchase Detail] where [AP POD Document No] = " & rsPurchase("AP PO Document No"), db, adOpenKeyset, adLockOptimistic, adCmdText
  '
  'rsDetail.MoveLast

  Dim Qty#
  Dim DetailUnitCost#
  Dim DetailItemCost#
  Dim CurrentLastCost#
  Dim CurrentAverageCost#
  Dim CurrentQuantityOnHand#
  Dim NewAverageCost#
  Dim Factor#
  Dim QtyRec#
  If rsDetail.RecordCount = 0 Then
    ' no detail data
  Else
    rsDetail.MoveFirst
    ' may have detail data
    Do While rsDetail.EOF = False
      'rsInventory.Index = "PrimaryKey"
      rsInventory.MoveFirst
      rsInventory.Find "[INV ITEM Id] ='" & rsDetail("AP POD Item ID") & "'"
      If rsInventory.EOF Then
        ' May be a non stock item
      Else
          ' Update Inventory Average & Last Cost
          Factor# = GetUOMMultiplier(rsDetail("AP POD Item ID"), rsDetail("AP POD Units"), db)
          DetailUnitCost# = IIf(IsNull(rsDetail("AP POD Unit Cost")), 0, rsDetail("AP POD Unit Cost"))
          
          If Factor# = 0 Then
          Else
            DetailItemCost# = DetailUnitCost# / Factor#
          End If

          If DetailUnitCost# > 0 Then
            CurrentLastCost# = IIf(IsNull(rsInventory("INV ITEM Last Cost")), 0, rsInventory("INV ITEM Last Cost"))
            CurrentAverageCost# = IIf(IsNull(rsInventory("INV ITEM Average Cost")), 0, rsInventory("INV ITEM Average Cost"))
            CurrentQuantityOnHand# = IIf(IsNull(rsInventory("INV ITEM Qty On Hand")), 0, rsInventory("INV ITEM Qty On Hand"))
            rsInventory("INV ITEM Last Cost") = DetailItemCost#
            QtyRec# = IIf(IsNull(rsDetail("AP POD Total Qty Received")), 0, rsDetail("AP POD Total Qty Received"))
            QtyRec# = QtyRec# * Factor#
            If CurrentQuantityOnHand# < 0 Then CurrentQuantityOnHand# = 0
            NewAverageCost# = ((CurrentQuantityOnHand# * CurrentAverageCost#) + (QtyRec# * DetailItemCost#))
            If CurrentQuantityOnHand# + QtyRec# <= 0 Then
              NewAverageCost# = DetailItemCost#
            Else
              NewAverageCost# = NewAverageCost# / (CurrentQuantityOnHand# + QtyRec#)
            End If
            rsInventory("INV ITEM Average Cost") = NewAverageCost#
          End If
          ' end of Update Inventory Average & Last Cost
          
          'Take units into consideration
          rsInventory("INV ITEM Qty On Hand") = rsInventory("INV ITEM Qty On Hand") + rsDetail("AP POD Total Qty Received") * Factor#
        rsInventory.Update
      End If
  
      rsDetail.MoveNext
    Loop
    rsInventory.Close
    Set rsInventory = Nothing
    
    rsDetail.Close
    Set rsDetail = Nothing

  End If
  ' End of Inventory Updates

  'Dim rsGLTrans As ADODB.Recordset
  'Set rsGLTrans = New ADODB.Recordset
  'rsGLTrans.Open "[GL Transaction]", db, adOpenStatic, adLockOptimistic, adCmdTable

  ' write GL Transaction Header
  Dim refr$
  Dim desc$
  Dim NewNumber&
      
  'db.Execute "DELETE * FROM [GL Transaction] WHERE [GL TRANS Document #]='" & "REC " & rsPurchase("AP PO Ext Document No") & "'"

  'rsGLTrans.AddNew
      SQLstatement = "INSERT INTO [GL Transaction]"
      SQLstatement = SQLstatement & " ([GL TRANS Document #],[GL TRANS Type],[GL TRANS Date],"
      SQLstatement = SQLstatement & " [GL TRANS Reference],[GL TRANS Amount],[GL TRANS Posted YN],"
      SQLstatement = SQLstatement & " [GL TRANS Description],[GL TRANS Source],[GL TRANS System Generated])"
    
    'rsGLTrans("GL TRANS Document #") = "REC " & rsPurchase("AP PO Ext Document No")
    Dim TempStr As String
    ' gl post date
    If PostDate% = 1 Then
      TempStr = Format(Now, "Short Date")
    Else
      TempStr = rsPurchase("AP PO Date")
    End If
    
    'rsGLTrans("GL TRANS Type") = "Receiving"

    SQLstatement = SQLstatement & " VALUES ('REC " & rsPurchase("AP PO Ext Document No") & "','Receiving',#" & TempStr & "#,"

    refr$ = rsPurchase("AP PO vendor Name")
    desc$ = IIf(IsNull(rsPurchase("AP PO Description")), "", rsPurchase("AP PO Description"))
    If Len(Trim$(desc$)) = 0 Then
      desc$ = "REC " & rsPurchase("AP PO Ext Document No")
    End If
      
      SQLstatement = SQLstatement & "'" & refr$ & "'," & rsPurchase("AP PO Total Amount") & ",1,"
      SQLstatement = SQLstatement & "'" & desc$ & "','REC " & rsPurchase("AP PO Ext Document No") & "',True)"
      'Debug.Print SQLstatement
      
      db.Execute SQLstatement
    
   ' rsGLTrans("GL TRANS Reference") = refr$
   ' rsGLTrans("GL TRANS Amount") = rsPurchase("AP PO Total Amount")
   ' rsGLTrans("GL TRANS Posted YN") = 1
   ' rsGLTrans("GL TRANS Description") = desc$
   ' rsGLTrans("GL TRANS Source") = "REC " & rsPurchase("AP PO Ext Document No")
   ' rsGLTrans("GL TRANS System Generated") = True
  'rsGLTrans.Update
  Dim rsGLTrans As ADODB.Recordset
  Set rsGLTrans = New ADODB.Recordset
  rsGLTrans.Open "SELECT [GL TRANS Number] FROM [GL Transaction] WHERE [GL TRANS Document #]='" & "REC " & rsPurchase("AP PO Ext Document No") & "'", db, adOpenKeyset, adLockReadOnly, adCmdText
      NewNumber& = rsGLTrans("GL TRANS Number")
  rsGLTrans.Close
  Set rsGLTrans = Nothing

  ' update GL
  '-----------------------------------------------------------------------
  ' PO with Payment GL Affected Accounts
  '
  '                  Debit   Credit   Source
  '                  -----   ------   ------
  ' Inventory          X              Item - Inventory
  ' Misc Charges       X              Pref - Purchases
  ' Freight Expense    X              Pref - Purchases
  ' Discount                   X      Pref - Purchases
  ' AP                         X      Pref - Purchases
  ' CASH                       X      Bank - Cash Acct
  ' The following entry is only valid if the payment exceeds the receipt total
  ' AP                 X              Pref - Purchases
  '
  ' Notes:
  ' Each Inventory Item is processed and the GL Acct is Retrieved.
  '-----------------------------------------------------------------------

  Dim InventoryAcct$

  ' Debits
  ' Inventory Increase
  Dim Longer&
  
  Longer& = 0
    
  Set rsDetail = New ADODB.Recordset
  rsDetail.Open "SELECT [AP POD Posting Account],[AP POD Item ID],[AP POD Item Total] " & _
  "FROM [AP Purchase Detail] where [AP POD Document No] = " & rsPurchase("AP PO Document No"), db, adOpenKeyset, adLockOptimistic, adCmdText
  'rsDetail.MoveLast
  'rsDetail.MoveFirst
'
  If rsDetail.RecordCount = 0 Then
    ' no detail data
  Else
    ' may have detail data
    rsDetail.MoveFirst
    Do While rsDetail.EOF = False
        
        ' use Account in detail 1st
        InventoryAcct$ = IIf(IsNull(rsDetail("AP POD Posting Account")), "", rsDetail("AP POD Posting Account"))

        ' use Account in Item 2nd
        If Len(InventoryAcct$) = 0 Then
          'rsInventory.Index = "PrimaryKey"
          'rsInventory.Seek rsDetail("AP POD Item ID")
          Set rsInventory = New ADODB.Recordset
          rsInventory.Open "SELECT [INV ITEM Inventory Account] FROM FROM [INV Items] WHERE [INV ITEM Id] ='" & rsDetail("AP POD Item ID") & "'", db, adOpenKeyset, adLockReadOnly, adCmdText
          'If rsInventory.BOF And rsInventory.EOF = True Then
            If rsInventory.RecordCount = 0 Then
            Else
              InventoryAcct$ = IIf(IsNull(rsInventory("INV ITEM Inventory Account")), "", rsInventory("INV ITEM Inventory Account"))
            End If
          rsInventory.Close
          Set rsInventory = Nothing
        End If

        ' use Account in Vendor 3rd
        If Len(InventoryAcct$) = 0 Then
          Dim VendorKey$
          
          VendorKey$ = IIf(IsNull(rsPurchase("AP PO Vendor ID")), "", rsPurchase("AP PO Vendor ID"))
          If Len(VendorKey$) > 0 Then
            'Set rsVendor = db2.OpenRecordset("AP Vendor")
            'rsVendor.Index = "PrimaryKey"
            'rsVendor.Seek "=", VendorKey$
            Set rsVendor = New ADODB.Recordset
            rsVendor.Open "SELECT [AP VEN Default GL] FROM [AP Vendor] WHERE [AP VEN ID]='" & VendorKey$ & "'", db, adOpenKeyset, adLockReadOnly, adCmdText
            'rsVendor.MoveFirst
            'rsVendor.Find "[AP VEN ID]='" & VendorKey$ & "'"
            If rsVendor.RecordCount > 0 Then
              InventoryAcct$ = IIf(IsNull(rsVendor("AP VEN Default GL")), "", rsVendor("AP VEN Default GL"))
            Else
              InventoryAcct$ = ""
            End If
            rsVendor.Close
            Set rsVendor = Nothing
            
          End If
        End If

        'Use preferences Acct 4th
        If Len(InventoryAcct$) = 0 Then InventoryAcct$ = rsCompany("SYS COM Purchase Inventory Acct")
        '+++
      SQLstatement = "INSERT INTO [GL Work Detail]"
      SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
      SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & InventoryAcct$ & "'," & rsDetail("AP POD Item Total") & ",0)"
      db.Execute SQLstatement
        
        'rsGLWorkDetail.AddNew
        '  rsGLWorkDetail("GW TRANSD Number") = NewNumber&
        '  rsGLWorkDetail("GW TRANSD Account") = InventoryAcct$
        '  rsGLWorkDetail("GW TRANSD Debit Amount") = rsDetail("AP POD Item Total")
        '  rsGLWorkDetail("GW TRANSD Credit Amount") = 0
        '  rsGLWorkDetail("GW TRANSD Project") = ""
        'rsGLWorkDetail.Update

      rsDetail.MoveNext
    Loop
  End If
  rsDetail.Close
  Set rsDetail = Nothing
  
  ' Freight Expense
  If rsPurchase("AP PO Shipping") > 0 Then
      SQLstatement = "INSERT INTO [GL Work Detail]"
      SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
      SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & rsCompany("SYS COM Purchase Freight Acct") & "'," & rsPurchase("AP PO Shipping") & ",0)"
      db.Execute SQLstatement
    'rsGLWorkDetail.AddNew
    '  rsGLWorkDetail("GW TRANSD Number") = NewNumber&
    '  rsGLWorkDetail("GW TRANSD Account") = rsCompany("SYS COM Purchase Freight Acct")
    '  rsGLWorkDetail("GW TRANSD Debit Amount") = rsPurchase("AP PO Shipping")
    '  rsGLWorkDetail("GW TRANSD Credit Amount") = 0
    '  rsGLWorkDetail("GW TRANSD Project") = ""
    'rsGLWorkDetail.Update
  End If

  ' Misc Charges
  If rsPurchase("AP PO Misc Charges") > 0 Then
      SQLstatement = "INSERT INTO [GL Work Detail]"
      SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
      SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & rsCompany("SYS COM Purchase Misc Acct") & "'," & rsPurchase("AP PO Misc Charges") & ",0)"
      db.Execute SQLstatement
    'rsGLWorkDetail.AddNew
    '  rsGLWorkDetail("GW TRANSD Number") = NewNumber&
    '  rsGLWorkDetail("GW TRANSD Account") = rsCompany("SYS COM Purchase Misc Acct")
    '  rsGLWorkDetail("GW TRANSD Debit Amount") = rsPurchase("AP PO Misc Charges")
    '  rsGLWorkDetail("GW TRANSD Credit Amount") = 0
    '  rsGLWorkDetail("GW TRANSD Project") = ""
    'rsGLWorkDetail.Update
  End If
  
  ' Credits
  ' Discount Amount
  If rsPurchase("AP PO Discount Amt") > 0 Then
      SQLstatement = "INSERT INTO [GL Work Detail]"
      SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
      SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & rsCompany("SYS COM Purchase Discount Acct") & "',0," & rsPurchase("AP PO Discount Amt") & ")"
      db.Execute SQLstatement
    'rsGLWorkDetail.AddNew
    '  rsGLWorkDetail("GW TRANSD Number") = NewNumber&
    '  rsGLWorkDetail("GW TRANSD Account") = rsCompany("SYS COM Purchase Discount Acct")
    '  rsGLWorkDetail("GW TRANSD Debit Amount") = 0
    '  rsGLWorkDetail("GW TRANSD Credit Amount") = rsPurchase("AP PO Discount Amt")
    '  rsGLWorkDetail("GW TRANSD Project") = ""
    'rsGLWorkDetail.Update
  End If
  
  ' Cash Payment
  If rsPurchase("AP PO Amount Paid") > 0 Then
      SQLstatement = "INSERT INTO [GL Work Detail]"
      SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
      SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & rsPurchase("AP PO Check Acct ID") & "',0," & rsPurchase("AP PO Amount Paid") & ")"
      db.Execute SQLstatement
    'rsGLWorkDetail.AddNew
    '  rsGLWorkDetail("GW TRANSD Number") = NewNumber&
    '  rsGLWorkDetail("GW TRANSD Account") = rsPurchase("AP PO Check Acct ID")
    '  rsGLWorkDetail("GW TRANSD Debit Amount") = 0
    '  rsGLWorkDetail("GW TRANSD Credit Amount") = rsPurchase("AP PO Amount Paid")
    '  rsGLWorkDetail("GW TRANSD Project") = ""
    'rsGLWorkDetail.Update

  End If

  ' AP
  If rsPurchase("AP PO Total Amount") - rsPurchase("AP PO Amount Paid") > 0 Then
      SQLstatement = "INSERT INTO [GL Work Detail]"
      SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
      SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & rsCompany("SYS COM Purchase AP Acct") & "',0," & rsPurchase("AP PO Total Amount") - rsPurchase("AP PO Amount Paid") & ")"
      db.Execute SQLstatement
    'rsGLWorkDetail.AddNew
    '  rsGLWorkDetail("GW TRANSD Number") = NewNumber&
    '  rsGLWorkDetail("GW TRANSD Account") = rsCompany("SYS COM Purchase AP Acct")
    '  rsGLWorkDetail("GW TRANSD Debit Amount") = 0
    '  rsGLWorkDetail("GW TRANSD Credit Amount") = rsPurchase("AP PO Total Amount") - rsPurchase("AP PO Amount Paid")
    '  rsGLWorkDetail("GW TRANSD Project") = ""
    'rsGLWorkDetail.Update
  End If


SkipIt2:


PostReceiving_Exit:

  ' post GL entry this receiving
  Dim Success%
  Success% = PostGLWorkDetail(TranDate, NewNumber&, db)
  If Success% = False Then
    MsgBox "An error occurred writing GL Transaction!", , "Error"
    PostReceiving = False
    Exit Function
  End If

  PostReceiving = True
  'On Error Resume Next
  'rsInventory.Close
  'Set rsInventory = Nothing
   rsCompany.Close
   Set rsCompany = Nothing
   rsPurchase.Close
   Set rsPurchase = Nothing
'   rsGLWorkDetail.Close
'   Set rsGLWorkDetail = Nothing
'   rsGLTrans.Close
'   Set rsGLTrans = Nothing
   'rsVendor.Close
   'Set rsVendor = Nothing
'   rsAPCross.Close
'   Set rsAPCross = Nothing
'   rsAPPaymentHeader.Close
'   Set rsAPPaymentHeader = Nothing
'   rsDetail.Close
'   Set rsDetail = Nothing
   'db.Close
   'Set db = Nothing
  Exit Function

PostReceiving_Error:
  Call ErrorLog("Purchase Module", "PostReceiving", Now, Err.Number, Err.Description, intShowError, db)
  PostReceiving = False
  Exit Function
'  Resume Next
  Exit Function

UnableToPostPOHere:
  PostReceiving = False
   'On Error Resume Next
   'rsInventory.Close
   'Set rsInventory = Nothing
   rsCompany.Close
   Set rsCompany = Nothing
   rsPurchase.Close
   Set rsPurchase = Nothing
'   rsGLWorkDetail.Close
'   Set rsGLWorkDetail = Nothing
'   rsGLTrans.Close
'   Set rsGLTrans = Nothing
   'rsVendor.Close
   'Set rsVendor = Nothing
'   rsAPCross.Close
'   Set rsAPCross = Nothing
'   rsAPPaymentHeader.Close
'   Set rsAPPaymentHeader = Nothing
'   rsDetail.Close
'   Set rsDetail = Nothing
   'db.Close
   'Set db = Nothing
  Exit Function

End Function

Function PostVoid(VendorID$, CheckNo$, db As ADODB.Connection) As Integer
'On Error GoTo PostVoid_Error
'Dim Currentdb As Boolean
'
  'Dim db As ADODB.Connection
  'Currentdb = False
  'If db Is Nothing Then
  '  Set db = New ADODB.Connection
  '  db.CursorLocation = adUseServer
  '  db.Open gblADOProvider
  '  Currentdb = True
  'End If
    
  Dim CurrentBalance@
  Dim msg$
  Dim title$

  Dim rsPayment As ADODB.Recordset
  Set rsPayment = New ADODB.Recordset
  rsPayment.Open "SELECT [AP PAY Transaction Date], [AP PAY Amount], [AP PAY ID], " & _
  "[AP PAY Bank Account] FROM [AP Payment Header] WHERE [AP PAY Check No]='" & _
  CheckNo$ & "' AND [AP PAY Vendor No]='" & VendorID$ & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
  'rsPayment.Index = "PrimaryKey"
  'rsPayment.Seek VendorID$, CheckNo$

  Dim rsCompany As ADODB.Recordset
  Set rsCompany = New ADODB.Recordset
  rsCompany.Open "SELECT  [SYS COM Purchase AP Acct],[SYS COM Purchase Discount Acct]," & _
  "[SYS COM GL Post By Date] FROM [SYS Company]", db, adOpenKeyset, adLockOptimistic, adCmdText
  rsCompany.MoveFirst
  
  'Post by 1 - system date or 2 - Transaction date?
  Dim PostDate%
  PostDate% = rsCompany("SYS COM GL Post By Date")

  'Set Invoice Type
  Dim TranDate As Variant

  'Set Post Date
  If PostDate% = 1 Then
    TranDate = DateValue(Format(Now, "Short Date"))
  Else
    TranDate = DateValue(rsPayment("AP PAY Transaction Date"))
  End If
  
  'Verify period can be posted to
  'Send TranDate
  'Return PeriodToPost and PeriodClosed
  Dim PeriodToPost%
  Dim PeriodClosed%
  
  Call VerifyPeriod(TranDate, PeriodToPost%, PeriodClosed%, db)
  
  If PeriodClosed% = True Then
    MsgBox "Unable to post transaction to a closed period.", , "Post Payment Error"
    PostVoid% = False
    Exit Function
  End If

  ' clear any GL Work records
  'Dim cmdtemp As ADODB.Recordset
  'Set cmdtemp = New ADODB.Recordset
  'cmdtemp.Open "DELETE DISTINCTROW * FROM [GL Work Detail]", db, , , adCmdText
  'cmdtemp.Close
  'Set cmdtemp = Nothing
  db.Execute "DELETE DISTINCTROW * FROM [GL Work Detail]", , adCmdText
  
  'Dim rsGLWorkDetail As ADODB.Recordset
  'Set rsGLWorkDetail = New ADODB.Recordset
  'rsGLWorkDetail.Open "[GL Work Detail]", db, adOpenStatic, adLockOptimistic, adCmdTable
  
  'Dim rsGLTrans As ADODB.Recordset
  'Set rsGLTrans = New ADODB.Recordset
  'rsGLTrans.Open "[GL Transaction]", db, adOpenStatic, adLockOptimistic, adCmdTable
  
  Dim rsPurchase As ADODB.Recordset
  Set rsPurchase = New ADODB.Recordset
  rsPurchase.Open "SELECT [AP PO Document No], [AP PO Balance Due], [AP PO Amount Paid] " & _
  "FROM [AP Purchase]", db, adOpenStatic, adLockOptimistic, adCmdText
  
  Dim TmpBalance@

  ' write GL Transaction Header
  Dim refr$
  Dim desc$
  Dim NewNumber&
  
  Dim rsVendor As ADODB.Recordset
  Set rsVendor = New ADODB.Recordset
  rsVendor.Open "SELECT [AP VEN ID], [AP VEN Name], [AP VEN Payments YTD], " & _
  "[AP VEN Payments Lifetime], [AP VEN Financial Period 1] FROM [AP Vendor] " & _
  "WHERE [AP VEN ID]='" & VendorID$ & "'", db, adOpenStatic, adLockOptimistic, adCmdText
  'rsVendor.Open "[AP Vendor]", db, adOpenStatic, adLockOptimistic, adCmdTable
  'rsVendor.Index = "PrimaryKey"
  'rsVendor.Seek VendorID$

  'rsGLTrans.AddNew
    
    'rsGLTrans("GL TRANS Document #") = "VOID " & CheckNo$
    'rsGLTrans("GL TRANS Type") = "Void Check"
    ' gl post date
      SQLstatement = "INSERT INTO [GL Transaction]"
      SQLstatement = SQLstatement & " ([GL TRANS Document #],[GL TRANS Type],[GL TRANS Date],"
      SQLstatement = SQLstatement & " [GL TRANS Reference],[GL TRANS Amount],[GL TRANS Posted YN],"
      SQLstatement = SQLstatement & " [GL TRANS Description],[GL TRANS Source],[GL TRANS System Generated])"
    
    Dim TempStr As String
    
    ' gl post date
    If PostDate% = 1 Then
      TempStr = Format(Now, "Short Date")
    Else
      TempStr = rsPayment("AP PAY Transaction Date")
    End If
      
      SQLstatement = SQLstatement & " VALUES ('VOID " & CheckNo$ & "','Void Check',#" & TempStr & "#,"
      
    If rsVendor.RecordCount = 0 Then
      refr$ = "Unknown"
    Else
      refr$ = rsVendor("AP VEN Name")
    End If

      SQLstatement = SQLstatement & "'" & refr$ & "'," & rsPayment("AP PAY Amount") & ",1,"
      SQLstatement = SQLstatement & "'CASH PAY " & CheckNo$ & "','VOID " & CheckNo$ & "',True)"
      'Debug.Print SQLstatement
      
      db.Execute SQLstatement
    
    'rsGLTrans("GL TRANS Reference") = refr$
    'rsGLTrans("GL TRANS Amount") = rsPayment("AP PAY Amount")
    'rsGLTrans("GL TRANS Posted YN") = 1
    'desc$ = "CASH PAY " & CheckNo$
    'rsGLTrans("GL TRANS Description") = desc$
    'rsGLTrans("GL TRANS Source") = "VOID " & CheckNo$
    'rsGLTrans("GL TRANS System Generated") = True
  'rsGLTrans.Update
  Dim rsGLTrans As ADODB.Recordset
  Set rsGLTrans = New ADODB.Recordset
  rsGLTrans.Open "SELECT [GL TRANS Number] FROM [GL Transaction]WHERE [GL TRANS Document #]='" & "VOID " & CheckNo$ & "'", db, adOpenKeyset, adLockReadOnly, adCmdText
      NewNumber& = rsGLTrans("GL TRANS Number")
  rsGLTrans.Close
  Set rsGLTrans = Nothing
  ' write GL Transaction Detail
'
  'Loop through line items
  Dim rsCross As ADODB.Recordset
  Set rsCross = New ADODB.Recordset
  rsCross.Open "SELECT [AP CROSS Applied Amount], [AP CROSS Discount Taken],[AP CROSS Write Off Amount]," & _
  "[AP CROSS Payed ID] FROM [AP Payment Invoice Cross Reference] WHERE " & _
  "[AP CROSS Payment ID] = " & rsPayment("AP PAY ID"), db, adOpenStatic, adLockOptimistic, adCmdText

  Dim TotalApplied@
  TotalApplied@ = 0
  If rsCross.RecordCount > 0 Then
    rsCross.MoveFirst
  End If
    Do While Not rsCross.EOF
      If rsCross("AP CROSS Applied Amount") > 0 Or rsCross("AP CROSS Discount Taken") > 0 Then
        ' process payments
        If rsCross("AP CROSS Applied Amount") > 0 Then
          
          ' update GL for payment
          ' write GL Transaction Detail
          '-----------------------------------------------------------------------
          ' Payment GL Affected Accounts
          '
          '                  Debit   Credit   Source
          '                  -----   ------   ------
          ' CASH               X              Bank - Cash Acct
          ' AP                         X      Pref - Sales
          '-----------------------------------------------------------------------

          ' Debits
      SQLstatement = "INSERT INTO [GL Work Detail]"
      SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
      SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & rsPayment("AP PAY Bank Account") & "" & "'," & rsCross("AP CROSS Applied Amount") & ",0)"
      db.Execute SQLstatement
      
      '    rsGLWorkDetail.AddNew
      '      rsGLWorkDetail("GW TRANSD Number") = 0
      '      rsGLWorkDetail("GW TRANSD Account") = rsPayment("AP PAY Bank Account")
      '      rsGLWorkDetail("GW TRANSD Debit Amount") = rsCross("AP CROSS Applied Amount")
      '      rsGLWorkDetail("GW TRANSD Credit Amount") = 0
      '      rsGLWorkDetail("GW TRANSD Project") = ""
      '    rsGLWorkDetail.Update
          
          ' Credits
          ' AP
      SQLstatement = "INSERT INTO [GL Work Detail]"
      SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
      SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & rsCompany("SYS COM Purchase AP Acct") & "" & "',0," & rsCross("AP CROSS Applied Amount") & ")"
      db.Execute SQLstatement
          
      '    rsGLWorkDetail.AddNew
      '      rsGLWorkDetail("GW TRANSD Number") = 0
      '      rsGLWorkDetail("GW TRANSD Account") = rsCompany("SYS COM Purchase AP Acct")
      '      rsGLWorkDetail("GW TRANSD Debit Amount") = 0
      '      rsGLWorkDetail("GW TRANSD Credit Amount") = rsCross("AP CROSS Applied Amount")
      '      rsGLWorkDetail("GW TRANSD Project") = ""
      '    rsGLWorkDetail.Update
          ' update GL for payment

        End If ' end process payments
        
        ' process discount amounts
        If rsCross("AP CROSS Discount Taken") > 0 Then

          ' update GL for discount
          '-----------------------------------------------------------------------
          ' Discount GL Affected Accounts
          '
          '                  Debit   Credit   Source
          '                  -----   ------   ------
          ' Discount           X               Pref - Sales
          ' AP                         X      Pref - Sales
          '-----------------------------------------------------------------------

          ' Debits
          ' Discount
      SQLstatement = "INSERT INTO [GL Work Detail]"
      SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
      SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & rsCompany("SYS COM Purchase Discount Acct") & "" & "'," & rsCross("AP CROSS Discount Taken") & ",0)"
      db.Execute SQLstatement
      
         'rsGLWorkDetail.AddNew
         '   rsGLWorkDetail("GW TRANSD Number") = 0
         '   rsGLWorkDetail("GW TRANSD Account") = rsCompany("SYS COM Purchase Discount Acct")
         '   rsGLWorkDetail("GW TRANSD Debit Amount") = rsCross("AP CROSS Discount Taken")
         '   rsGLWorkDetail("GW TRANSD Credit Amount") = 0
         '   rsGLWorkDetail("GW TRANSD Project") = ""
         ' rsGLWorkDetail.Update

          ' Credits
          ' AP
      SQLstatement = "INSERT INTO [GL Work Detail]"
      SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
      SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & rsCompany("SYS COM Purchase Discount Acct") & "" & "',0," & rsCross("AP CROSS Discount Taken") & ")"
      db.Execute SQLstatement
      
      '    rsGLWorkDetail.AddNew
      '      rsGLWorkDetail("GW TRANSD Number") = 0
      '      rsGLWorkDetail("GW TRANSD Account") = rsCompany("SYS COM Purchase AP Acct")
      '      rsGLWorkDetail("GW TRANSD Debit Amount") = 0
      '      rsGLWorkDetail("GW TRANSD Credit Amount") = rsCross("AP CROSS Discount Taken")
      '      rsGLWorkDetail("GW TRANSD Project") = ""
      '    rsGLWorkDetail.Update
          ' update GL for discount

        End If ' end process discount amounts
        
        ' update AP OPEN Payment
        'rsPurchase.Index = "PrimaryKey"
        'rsPurchase.Seek rsCross("AP CROSS Payed ID")'
        rsPurchase.MoveFirst
        rsPurchase.Find "[AP PO Document No]=" & rsCross("AP CROSS Payed ID")
          TmpBalance@ = IIf(IsNull(rsPurchase("AP PO Balance Due")), 0, rsPurchase("AP PO Balance Due"))
          
          If rsCross("AP CROSS Applied Amount") <> 0 Then
            TmpBalance@ = TmpBalance + rsCross("AP CROSS Applied Amount")
          End If
          If rsCross("AP CROSS Discount Taken") <> 0 Then
            TmpBalance@ = TmpBalance@ + rsCross("AP CROSS Discount Taken")
          End If
          rsPurchase("AP PO Balance Due") = TmpBalance@
          rsPurchase("AP PO Amount Paid") = rsPurchase("AP PO Amount Paid") - rsCross("AP CROSS Applied Amount") - rsCross("AP CROSS Discount Taken")

        rsPurchase.Update
        ' end of update AP OPEN Payment

      Else
      End If
    rsCross.MoveNext
  Loop

  ' update Vendor stats
  'rsVendor.MoveFirst
  'rsVendor.Index = "PrimaryKey"
  'rsVendor.Seek VendorID$
  rsVendor.MoveFirst
  'rsVendor.Find "[AP VEN ID]='" & VendorID$ & "'"

    rsVendor("AP VEN Payments YTD") = rsVendor("AP VEN Payments YTD") - rsPayment("AP PAY Amount")
    rsVendor("AP VEN Payments Lifetime") = rsVendor("AP VEN Payments Lifetime") - rsPayment("AP PAY Amount")
    ' Update current Balance - if not Paid in full
    CurrentBalance@ = IIf(IsNull(rsVendor("AP VEN Financial Period 1")), 0, rsVendor("AP VEN Financial Period 1"))
    CurrentBalance@ = CurrentBalance@ + rsPayment("AP PAY Amount")
    rsVendor("AP VEN Financial Period 1") = CurrentBalance@
  rsVendor.Update
  
  'post GL Entry
  Dim Success%
  Success% = PostGLWorkDetail(TranDate, NewNumber&, db)

  PostVoid = True
  
PostVoid_Exit:

  'On Error Resume Next
   rsCompany.Close
   Set rsCompany = Nothing
   rsPurchase.Close
   Set rsPurchase = Nothing
   'rsGLWorkDetail.Close
   'Set rsGLWorkDetail = Nothing
   'rsGLTrans.Close
   'Set rsGLTrans = Nothing
   rsVendor.Close
   Set rsVendor = Nothing
   rsCross.Close
   Set rsCross = Nothing
   rsPayment.Close
   Set rsPayment = Nothing
   'rsDetail.Close
   'Set rsDetail = Nothing
   'If Currentdb = True Then
   ' db.Close
   ' Set db = Nothing
   'End If
  Exit Function

PostVoid_Error:
  Call ErrorLog("Purchase Module", "PostVoid", Now, Err.Number, Err.Description, True, db)
  PostVoid = False
  Resume Next

End Function

Function PostVoucher(DocumentKey&, intShowError As Integer, db As ADODB.Connection)

  'On Error GoTo PostVoucher_Error
  
  'Dim db As ADODB.Connection
  'Set db = New ADODB.Connection
  'db.CursorLocation = adUseServer
  'db.Open gblADOProvider
  
  Dim msg$
  Dim title$

  Dim rsCompany As ADODB.Recordset
  Set rsCompany = New ADODB.Recordset
  'rsCompany.Open "[SYS Company]", db, adOpenStatic, adLockOptimistic, adCmdTable
  rsCompany.Open "SELECT  [SYS COM Purchase AP Acct],[SYS COM Purchase Discount Acct]," & _
  "[SYS COM GL Post By Date],[SYS COM Purchase Inventory Acct],[SYS COM Purchase Freight Acct]," & _
  "[SYS COM Purchase Misc Acct] FROM [SYS Company]", db, adOpenKeyset, adLockOptimistic, admcdtext
  rsCompany.MoveFirst

  Dim rsPurchase As ADODB.Recordset
  Set rsPurchase = New ADODB.Recordset
  rsPurchase.Open "SELECT [AP PO Document Type],[AP PO Date],[AP PO Vendor ID],[AP PO Amount Paid]," & _
  "[AP PO Payment Method],[AP PO Check Number],[AP PO Check Acct ID],[AP PO Document No]," & _
  "[AP PO Ext Document No],[AP PO vendor Name],[AP PO Description],[AP PO Total Amount]," & _
  "[AP PO Shipping],[AP PO Misc Charges],[AP PO Discount Amt] FROM [AP Purchase] " & _
  "WHERE [AP PO Document No]=" & DocumentKey&, db, adOpenKeyset, adLockOptimistic, adCmdText
  
  ' first lets get the Voucher
  'rsPurchase.Index = "PrimaryKey"
  'rsPurchase.Seek DocumentKey&
  'rsPurchase.MoveFirst
  'rsPurchase.Find "[AP PO Document No]=" & DocumentKey&

  'Post by 1 - system date or 2 - Transaction date?
  Dim PostDate%
  PostDate% = rsCompany("SYS COM GL Post By Date")

  'Set Purchase Type
  Dim PurchaseType$
  PurchaseType$ = rsPurchase("AP PO Document Type")

  'Set Post Date
  Dim TranDate As Variant
  If PostDate% = 1 Then
    TranDate = DateValue(Format(Now, "Short Date"))
  Else
    TranDate = DateValue(rsPurchase("AP PO Date"))
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
    GoTo UnableToPostVoucherHere
  End If
  
  'On Error GoTo PostVoucher_Error

  ' clear any GL Work records
  'Dim cmdtemp As ADODB.Recordset
  'Set cmdtemp = New ADODB.Recordset
  'cmdtemp.Open "DELETE DISTINCTROW * FROM [GL Work Detail]", db, , , adCmdText
  'cmdtemp.Close
  'Set cmdtemp = Nothing
  db.Execute "DELETE DISTINCTROW * FROM [GL Work Detail]"

  'Dim rsGLWorkDetail As ADODB.Recordset
  'Set rsGLWorkDetail = New ADODB.Recordset
  'rsGLWorkDetail.Open "[GL Work Detail]", db, adOpenStatic, adLockOptimistic, adCmdTable
  
  ' update vendor stats
  Dim rsVendor As ADODB.Recordset
  Set rsVendor = New ADODB.Recordset
  rsVendor.Open "SELECT [AP VEN ID], [AP VEN Name], [AP VEN Payments YTD]," & _
  "[AP VEN Payment Number Lifetime],[AP VEN Payment Number YTD],[AP VEN Purchase YTD], " & _
  "[AP VEN Purchase Lifetime],[AP VEN Purchase Number Lifetime],[AP VEN Purchase Number YTD]," & _
  "[AP VEN Payments Lifetime], [AP VEN Financial Period 1] FROM [AP Vendor] " & _
  "WHERE [AP VEN ID]='" & rsPurchase("AP PO Vendor ID") & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
  'rsVendor.Open "SELECT * FROM [AP Vendor] where [AP VEN ID] = '" & rsPurchase("AP PO Vendor ID") & "'", db, adOpenStatic, adLockOptimistic, adCmdText
  
  Dim CurrentBalance@

    If rsPurchase("AP PO Amount Paid") > 0 Then
      rsVendor("AP VEN Payments YTD") = rsVendor("AP VEN Payments YTD") + rsPurchase("AP PO Amount Paid")
      rsVendor("AP VEN Payments Lifetime") = rsVendor("AP VEN Payments Lifetime") + rsPurchase("AP PO Amount Paid")
      rsVendor("AP VEN Payment Number Lifetime") = rsVendor("AP VEN Payment Number Lifetime") + 1
      rsVendor("AP VEN Payment Number YTD") = rsVendor("AP VEN Payment Number YTD") + 1
    End If
    rsVendor("AP VEN Purchase YTD") = rsVendor("AP VEN Purchase YTD") + rsPurchase("AP PO Total Amount")
    rsVendor("AP VEN Purchase Lifetime") = rsVendor("AP VEN Purchase Lifetime") + rsPurchase("AP PO Total Amount")
    rsVendor("AP VEN Purchase Number Lifetime") = rsVendor("AP VEN Purchase Number Lifetime") + 1
    rsVendor("AP VEN Purchase Number YTD") = rsVendor("AP VEN Purchase Number YTD") + 1
    ' Update current Balance - if not Paid in full
    If rsPurchase("AP PO Total Amount") - rsPurchase("AP PO Amount Paid") > 0 Then
      CurrentBalance@ = IIf(IsNull(rsVendor("AP VEN Financial Period 1")), 0, rsVendor("AP VEN Financial Period 1"))
      CurrentBalance@ = CurrentBalance@ + (rsPurchase("AP PO Total Amount") - rsPurchase("AP PO Amount Paid"))
      rsVendor("AP VEN Financial Period 1") = CurrentBalance@
    End If
  rsVendor.Update
  rsVendor.Close
  Set rsVendor = Nothing

  Dim PaymentID&

  If rsPurchase("AP PO Amount Paid") > 0 Then
    
    'Dim rsAPPaymentHeader As ADODB.Recordset
    'Set rsAPPaymentHeader = New ADODB.Recordset
    'rsAPPaymentHeader.Open "[AP Payment Header]", db, adOpenStatic, adLockOptimistic, adCmdTable
    
    'Dim rsAPCross As ADODB.Recordset
    'Set rsAPCross = New ADODB.Recordset
    'rsAPCross.Open "[AP Payment Invoice Cross Reference]", db, adOpenStatic, adLockOptimistic, adCmdTable
    
    SQLstatement = "INSERT INTO [AP Payment Header]"
    SQLstatement = SQLstatement & " ([AP PAY Type],[AP PAY Check No],[AP PAY Vendor No]," & _
    "[AP PAY Transaction Date],[AP PAY Amount],[AP PAY UnApplied Amount]," & _
    "[AP PAY Bank Account],[AP PAY Status],[AP PAY Posted YN],[AP PAY Cleared],[AP PAY Void])"
    
    '  write Payment Header
    Dim PayType As String
    'rsAPPaymentHeader.AddNew
      If rsPurchase("AP PO Payment Method") = "Cash" Or rsPurchase("AP PO Payment Method") = "Company Check" Then
        PayType = "Payment"
      Else
        PayType = "Charge"
      End If
    '  rsAPPaymentHeader("AP PAY Check No") = rsPurchase("AP PO Check Number")
    '  rsAPPaymentHeader("AP PAY Vendor No") = rsPurchase("AP PO Vendor ID") & ""
    '  rsAPPaymentHeader("AP PAY Transaction Date") = rsPurchase("AP PO Date")
    '  rsAPPaymentHeader("AP PAY Amount") = rsPurchase("AP PO Amount Paid")
    '  rsAPPaymentHeader("AP PAY UnApplied Amount") = 0
    '  rsAPPaymentHeader("AP PAY Bank Account") = rsPurchase("AP PO Check Acct ID")
    '  rsAPPaymentHeader("AP PAY Status") = "Posted"
    '  rsAPPaymentHeader("AP PAY Posted YN") = True
    '  rsAPPaymentHeader("AP PAY Cleared") = False
    '  rsAPPaymentHeader("AP PAY Void") = False
    'rsAPPaymentHeader.Update
    SQLstatement = SQLstatement & " VALUES ('" & PayType & "','" & _
    rsPurchase("AP PO Check Number") & "','" & rsPurchase("AP PO Vendor ID") & "',#" & _
    rsPurchase("AP PO Date") & "#," & rsPurchase("AP PO Amount Paid") & ",0,'" & _
    rsPurchase("AP PO Check Acct ID") & "','Posted',True,False,False)"
    db.Execute SQLstatement
    
      Dim rsAPPaymentHeader As ADODB.Recordset
      Set rsAPPaymentHeader = New ADODB.Recordset
      rsAPPaymentHeader.Open "SELECT [AP PAY ID] FROM [AP Payment Header] WHERE [AP PAY Check No]='" & rsPurchase("AP PO Check Number") & "'", db, adOpenKeyset, adLockReadOnly, adCmdText
      If rsAPPaymentHeader.RecordCount > 1 Then
        rsAPPaymentHeader.MoveLast
        PaymentID& = rsAPPaymentHeader("AP PAY ID")
      Else
        PaymentID& = rsAPPaymentHeader("AP PAY ID")
      End If
      rsAPPaymentHeader.Close
      Set rsAPPaymentHeader = Nothing
    ' end of write payment header
  
    Dim AmountPaid As Currency
    ' write AR Cross reference Record
      SQLstatement = "INSERT INTO [AP Payment Invoice Cross Reference]"
      SQLstatement = SQLstatement & " ([AP CROSS Payment ID],[AP CROSS Payed ID],"
      SQLstatement = SQLstatement & " [AP CROSS Discount Taken],[AP CROSS Write Off Amount]"
      SQLstatement = SQLstatement & " [AP CROSS Applied Amount],[AP CROSS Cleared])"
      SQLstatement = SQLstatement & " VALUES (" & PaymentID& & ",'" & rsPurchase("AP PO Document No") & "',0,0,"
      
      AmountPaid = IIf(IsNull(rsPurchase("AP PO Amount Paid")), 0, rsPurchase("AP PO Amount Paid"))
      SQLstatement = SQLstatement & AmountPaid & ",True)"
      db.Execute SQLstatement
    'rsAPCross.AddNew
    '  rsAPCross("AP CROSS Payment ID") = PaymentID&
    '  rsAPCross("AP CROSS Payed ID") = rsPurchase("AP PO Document No")
    '  rsAPCross("AP CROSS Discount Taken") = 0
    '  rsAPCross("AP CROSS Write Off Amount") = 0
    '  rsAPCross("AP CROSS Applied Amount") = IIf(IsNull(rsPurchase("AP PO Amount Paid")), 0, rsPurchase("AP PO Amount Paid"))
    '  rsAPCross("AP CROSS Cleared") = True
    'rsAPCross.Update
    ' end of write payment header
  End If

  ' End of New AP Cross Payment and AP Payment Header
  '--------------------------------------------------

  'Dim rsGLTrans As ADODB.Recordset
  'Set rsGLTrans = New ADODB.Recordset
  'rsGLTrans.Open "[GL Transaction]", db, adOpenStatic, adLockOptimistic, adCmdTable

  ' write GL Transaction Header
  Dim refr$
  Dim desc$
  Dim NewNumber&
  'db.Execute "DELETE * FROM [GL Transaction] WHERE [GL TRANS Document #]='" & "VOU " & rsPurchase("AP PO Ext Document No") & "'"
  SQLstatement = "INSERT INTO [GL Transaction]"
  SQLstatement = SQLstatement & " ([GL TRANS Document #],[GL TRANS Type],[GL TRANS Date],"
  SQLstatement = SQLstatement & " [GL TRANS Reference],[GL TRANS Amount],[GL TRANS Posted YN],"
  SQLstatement = SQLstatement & " [GL TRANS Description],[GL TRANS Source],[GL TRANS System Generated])"
  'rsGLTrans.AddNew
    
  '  rsGLTrans("GL TRANS Document #") = "VOU " & rsPurchase("AP PO Ext Document No")
    
    ' gl post date
    Dim TempStr As String
    If PostDate% = 1 Then
      TempStr = Format(Now, "Short Date")
    Else
      TempStr = rsPurchase("AP PO Date")
    End If
    
    'rsGLTrans("GL TRANS Type") = "Voucher"
    
    SQLstatement = SQLstatement & " VALUES ('VOU " & rsPurchase("AP PO Ext Document No") & "','Voucher',#" & TempStr & "#,"

    refr$ = rsPurchase("AP PO vendor Name")
    desc$ = IIf(IsNull(rsPurchase("AP PO Description")), "", rsPurchase("AP PO Description"))
    If Len(Trim$(desc$)) = 0 Then
      desc$ = "VOU " & rsPurchase("AP PO Ext Document No")
    End If
    
      SQLstatement = SQLstatement & "'" & refr$ & "'," & rsPurchase("AP PO Total Amount") & ",1,"
      SQLstatement = SQLstatement & "'" & desc$ & "','VOU " & rsPurchase("AP PO Ext Document No") & "',True)"
      'Debug.Print SQLstatement
      
      db.Execute SQLstatement
    
  '  rsGLTrans("GL TRANS Reference") = refr$
  '  rsGLTrans("GL TRANS Amount") = rsPurchase("AP PO Total Amount")
  '  rsGLTrans("GL TRANS Posted YN") = 1
  '  rsGLTrans("GL TRANS Description") = desc$
  '  rsGLTrans("GL TRANS Source") = "VOU " & rsPurchase("AP PO Ext Document No")
  '  rsGLTrans("GL TRANS System Generated") = True
  'rsGLTrans.Update
  Dim rsGLTrans As ADODB.Recordset
  Set rsGLTrans = New ADODB.Recordset
  rsGLTrans.Open "SELECT [GL TRANS Number] FROM [GL Transaction] WHERE [GL TRANS Document #]='" & "VOU " & rsPurchase("AP PO Ext Document No") & "'", db, adOpenKeyset, adLockReadOnly, adCmdText
      NewNumber& = rsGLTrans("GL TRANS Number")
  rsGLTrans.Close
  Set rsGLTrans = Nothing
  
  ' update GL
  '-----------------------------------------------------------------------
  ' Voucher with Payment GL Affected Accounts
  '
  '                  Debit   Credit   Source
  '                  -----   ------   ------
  ' Voucher Item       X              AP Detail
  ' AP                         X      Pref - Purchases
  ' CASH                       X      Bank - Cash Acct
  ' Discount                   X      Pref - PO Discount Acct xxx 10/23/95
  ' ?Misc Charges       X              Pref - PO Misc Charges Acct
  ' ?Freight Expenses   X              Pref - PO Freight Expenses
  ' The following entry is only valid if the payment exceeds the voucher total
  ' AP                 X              Pref - Purchases
  '
  ' Notes:
  ' Each Detail is processed and the GL Acct is Retrieved.
  '-----------------------------------------------------------------------

  Dim APDetailAcct$

  ' Debits
  ' Detail Item Increase
  Dim Longer&
  Longer& = 0
  Dim rsDetail As ADODB.Recordset
  Set rsDetail = New ADODB.Recordset
  rsDetail.Open "SELECT [AP POD Posting Account],[AP POD Item ID],[AP POD Item Total] " & _
  "FROM [AP Purchase Detail] where [AP POD Document No] = " & rsPurchase("AP PO Document No"), db, adOpenKeyset, adLockOptimistic, adCmdText
  'rsDetail.MoveLast
  'rsDetail.MoveFirst

  If rsDetail.RecordCount = 0 Then
    ' no detail data
  Else
    ' may have detail data
    rsDetail.MoveFirst
    Do While rsDetail.EOF = False
      ' use Account in detail 1st
      APDetailAcct$ = IIf(IsNull(rsDetail("AP POD Posting Account")), "", rsDetail("AP POD Posting Account"))

      ' use Account in Vendor 2nd
      If Len(APDetailAcct$) = 0 Then
        Dim VendorKey$
        VendorKey$ = IIf(IsNull(rsPurchase("AP PO Vendor ID")), "", rsPurchase("AP PO Vendor ID"))
        If Len(VendorKey$) > 0 Then
          Set rsVendor = New ADODB.Recordset
          rsVendor.Open "SELECT [AP VEN Default GL] FROM FROM [AP Vendor] WHERE [AP VEN ID] ='" & VendorKey$ & "'", db, adOpenKeyset, adLockReadOnly, adCmdText
          'Set rsVendor = db2.OpenRecordset("AP Vendor")
          'rsVendor.Index = "PrimaryKey"
          'rsVendor.Seek "=", VendorKey$
          If rsVendor.RecordCount = 0 Then
          Else
            APDetailAcct$ = IIf(IsNull(rsVendor("AP VEN Default GL")), "", rsVendor("AP VEN Default GL"))
          End If
          rsVendor.Close
          Set rsVendor = Nothing
        End If
      End If

      'Use Misc Purchase Acct 3rd
      If Len(APDetailAcct$) = 0 Then APDetailAcct$ = rsCompany("SYS COM Purchase Misc Acct")
      SQLstatement = "INSERT INTO [GL Work Detail]"
      SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
      SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & APDetailAcct$ & "'," & rsDetail("AP POD Item Total") & ",0)"
      db.Execute SQLstatement
      'rsGLWorkDetail.AddNew
      '  rsGLWorkDetail("GW TRANSD Number") = NewNumber&
      '  rsGLWorkDetail("GW TRANSD Account") = APDetailAcct$
      '  rsGLWorkDetail("GW TRANSD Debit Amount") = rsDetail("AP POD Item Total")
      '  rsGLWorkDetail("GW TRANSD Credit Amount") = 0
      '  rsGLWorkDetail("GW TRANSD Project") = ""
      'rsGLWorkDetail.Update
      
      rsDetail.MoveNext
    Loop
  End If
  rsDetail.Close
  Set rsDetail = Nothing
  
  'Freight Amount
  If rsPurchase("AP PO Shipping") > 0 Then
      SQLstatement = "INSERT INTO [GL Work Detail]"
      SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
      SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & rsCompany("SYS COM Purchase Freight Acct") & "'," & rsPurchase("AP PO Shipping") & ",0)"
      db.Execute SQLstatement
    'rsGLWorkDetail.AddNew
    '  rsGLWorkDetail("GW TRANSD Number") = NewNumber&
    '  rsGLWorkDetail("GW TRANSD Account") = rsCompany("SYS COM Purchase Freight Acct")
    '  rsGLWorkDetail("GW TRANSD Debit Amount") = rsPurchase("AP PO Shipping")
    '  rsGLWorkDetail("GW TRANSD Credit Amount") = 0
    '  rsGLWorkDetail("GW TRANSD Project") = ""
    'rsGLWorkDetail.Update
  End If

  'Misc Charges Amount
  If rsPurchase("AP PO Misc Charges") > 0 Then
      SQLstatement = "INSERT INTO [GL Work Detail]"
      SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
      SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & rsCompany("SYS COM Purchase Misc Acct") & "'," & rsPurchase("AP PO Misc Charges") & ",0)"
      db.Execute SQLstatement
    'rsGLWorkDetail.AddNew
    '  rsGLWorkDetail("GW TRANSD Number") = NewNumber&
    '  rsGLWorkDetail("GW TRANSD Account") = rsCompany("SYS COM Purchase Misc Acct")
    '  rsGLWorkDetail("GW TRANSD Debit Amount") = rsPurchase("AP PO Misc Charges")
    '  rsGLWorkDetail("GW TRANSD Credit Amount") = 0
    '  rsGLWorkDetail("GW TRANSD Project") = ""
    'rsGLWorkDetail.Update
  End If

  ' Credits
  ' Cash Payment
  If rsPurchase("AP PO Amount Paid") > 0 Then
      SQLstatement = "INSERT INTO [GL Work Detail]"
      SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
      SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & rsPurchase("AP PO Check Acct ID") & "',0," & rsPurchase("AP PO Amount Paid") & ")"
      db.Execute SQLstatement
    'rsGLWorkDetail.AddNew
    '  rsGLWorkDetail("GW TRANSD Number") = NewNumber&
    '  rsGLWorkDetail("GW TRANSD Account") = rsPurchase("AP PO Check Acct ID")
    '  rsGLWorkDetail("GW TRANSD Debit Amount") = 0
    '  rsGLWorkDetail("GW TRANSD Credit Amount") = rsPurchase("AP PO Amount Paid")
    '  rsGLWorkDetail("GW TRANSD Project") = ""
    'rsGLWorkDetail.Update
  End If
    
  'Discount Amount
  If rsPurchase("AP PO Discount Amt") > 0 Then
      SQLstatement = "INSERT INTO [GL Work Detail]"
      SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
      SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & rsCompany("SYS COM Purchase Discount Acct") & "',0," & rsPurchase("AP PO Discount Amt") & ")"
      db.Execute SQLstatement
    'rsGLWorkDetail.AddNew
    '  rsGLWorkDetail("GW TRANSD Number") = NewNumber&
    '  rsGLWorkDetail("GW TRANSD Account") = rsCompany("SYS COM Purchase Discount Acct")
    '  rsGLWorkDetail("GW TRANSD Debit Amount") = 0
    '  rsGLWorkDetail("GW TRANSD Credit Amount") = rsPurchase("AP PO Discount Amt")
    '  rsGLWorkDetail("GW TRANSD Project") = ""
    'rsGLWorkDetail.Update
  End If

  ' AP
  If rsPurchase("AP PO Total Amount") - rsPurchase("AP PO Amount Paid") > 0 Then
      SQLstatement = "INSERT INTO [GL Work Detail]"
      SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
      SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & rsCompany("SYS COM Purchase AP Acct") & "',0," & rsPurchase("AP PO Total Amount") - rsPurchase("AP PO Amount Paid") & ")"
      db.Execute SQLstatement
    'rsGLWorkDetail.AddNew
    '  rsGLWorkDetail("GW TRANSD Number") = NewNumber&
    '  rsGLWorkDetail("GW TRANSD Account") = rsCompany("SYS COM Purchase AP Acct")
    '  rsGLWorkDetail("GW TRANSD Debit Amount") = 0
    '  rsGLWorkDetail("GW TRANSD Credit Amount") = rsPurchase("AP PO Total Amount") - rsPurchase("AP PO Amount Paid")
    '  rsGLWorkDetail("GW TRANSD Project") = ""
    'rsGLWorkDetail.Update
  End If


  ' post GL entry for this Voucher
  Dim Success%
  Success% = PostGLWorkDetail(TranDate, NewNumber&, db)
  If Success% = False Then
    MsgBox "An error occurred writing GL Transaction!", , "Error"
    PostVoucher = False
    GoTo PostVoucher_Exit:
  End If

  PostVoucher = True

PostVoucher_Exit:
   'On Error Resume Next
   rsCompany.Close
   Set rsCompany = Nothing
   rsPurchase.Close
   Set rsPurchase = Nothing
   'rsGLWorkDetail.Close
   'Set rsGLWorkDetail = Nothing
   'rsGLTrans.Close
   'Set rsGLTrans = Nothing
   'rsVendor.Close
   'Set rsVendor = Nothing
'   rsCross.Close
   'Set rsCross = Nothing
'   rsPayment.Close
   'Set rsPayment = Nothing
   'rsDetail.Close
   'Set rsDetail = Nothing
   'db.Close
   'Set db = Nothing
  Exit Function

PostVoucher_Error:
  Call ErrorLog("Purchase Module", "PostVoucher", Now, Err.Number, Err.Description, intShowError, db)
  PostVoucher = False
  'On Error Resume Next
   rsCompany.Close
   Set rsCompany = Nothing
   rsPurchase.Close
   Set rsPurchase = Nothing
   'rsGLWorkDetail.Close
   'Set rsGLWorkDetail = Nothing
   'rsGLTrans.Close
   'Set rsGLTrans = Nothing
   'rsVendor.Close
   'Set rsVendor = Nothing
   'rsCross.Close
   'Set rsCross = Nothing
   'rsPayment.Close
   'Set rsPayment = Nothing
   'rsDetail.Close
   'Set rsDetail = Nothing
   'db.Close
   'Set db = Nothing
  Exit Function

UnableToPostVoucherHere:
  PostVoucher = False
  'On Error Resume Next
   rsCompany.Close
   Set rsCompany = Nothing
   rsPurchase.Close
   Set rsPurchase = Nothing
   rsGLWorkDetail.Close
   'Set rsGLWorkDetail = Nothing
   'rsGLTrans.Close
   'Set rsGLTrans = Nothing
   'rsVendor.Close
   'Set rsVendor = Nothing
   'rsCross.Close
   'Set rsCross = Nothing
   'rsPayment.Close
   'Set rsPayment = Nothing
   'rsDetail.Close
   'Set rsDetail = Nothing
   'db.Close
   'Set db = Nothing
  Exit Function

End Function

Function PrintCheck(VendorID$, CheckNo$, BankID$, TDate As Variant, Optional db As ADODB.Connection) As Integer
Dim Currentdb As Boolean
  
  'Dim db As ADODB.Connection
  Currentdb = False
  If db Is Nothing Then
    Set db = New ADODB.Connection
    db.CursorLocation = adUseServer
    db.Open gblADOProvider
    Currentdb = True
  End If
 
  Dim rsAPHeader As ADODB.Recordset
  Dim rsAPDetail As ADODB.Recordset
  Dim rsWork As ADODB.Recordset

'  Dim cmdtemp As ADODB.Recordset
'  Set cmdtemp = New ADODB.Recordset
'  cmdtemp.Open "DELETE DISTINCTROW * FROM [Print Check Work]", db, , , adCmdText
  'cmdtemp.Close
'  Set cmdtemp = Nothing

  db.Execute "DELETE DISTINCTROW * FROM [Print Check Work]", , adCmdText

  Set rsWork = New ADODB.Recordset
  rsWork.Open "[Print Check Work]", db, adOpenKeyset, adLockOptimistic, adCmdTable

  Set rsAPHeader = New ADODB.Recordset
  rsAPHeader.Open "SELECT [AP PAY ID],[AP PAY Amount] FROM [AP Payment Header] WHERE [AP PAY Check No] = '" & CheckNo$ & "' AND [AP PAY Vendor No] = '" & VendorID$ & "' AND [AP PAY Bank Account] = '" & BankID$ & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
  
  'If rsAPHeader.BOF And rsAPHeader.EOF Then
  If rsAPHeader.RecordCount = 0 Then
    'This is a problem
    PrintCheck% = False
    GoTo NoChecks:
  End If

  Dim X%
  
  X% = 0

  Dim TotalAmount@

  TotalAmount@ = NZ(SumRecord("[AP CROSS Applied Amount]", "[Work - Payment Information]", db, "[AP CROSS Payment ID] = " & rsAPHeader("AP PAY ID")))

  If TotalAmount@ = 0 Then
    rsWork.AddNew
      rsWork("Vendor ID") = VendorID$
      rsWork("Check Number") = CheckNo$
      rsWork("Total Amount") = rsAPHeader("AP PAY Amount")
      rsWork("Transaction Date") = TDate
      rsWork("Visible") = False
      rsWork("Order") = -1
    rsWork.Update
  Else
    Set rsAPDetail = New ADODB.Recordset
    rsAPDetail.Open "SELECT * FROM [Work - Payment Information] where [AP CROSS Payment ID] = " & rsAPHeader("AP PAY ID") & " ORDER BY [AP PO Date] Desc", db, adOpenStatic, adLockOptimistic, adCmdText
    'On Error Resume Next
    rsAPDetail.MoveFirst
    Do While Not rsAPDetail.EOF
      rsWork.AddNew
        rsWork("Vendor ID") = VendorID$
        rsWork("Check Number") = CheckNo$
        rsWork("Total Amount") = rsAPHeader("AP PAY Amount")
        rsWork("Transaction Date") = TDate
        If X% > 10 Then
          rsWork("Visible") = False
          rsWork("Order") = -1
        Else
          rsWork("Visible") = True
          rsWork("Order") = 11 - X%
        End If
        rsWork("Reference #") = rsAPDetail("AP PO Ext Document No") & ""
        If IsNull(rsAPDetail("AP PO Vendor Invoice No")) Then
        Else
            rsWork("Invoice #") = rsAPDetail("AP PO Vendor Invoice No") & ""
        End If
        rsWork("Invoice Date") = rsAPDetail("AP PO Date") & ""
        rsWork("Invoice Amt") = rsAPDetail("AP PO Total Amount")
        rsWork("Amount Paid") = rsAPDetail("AP CROSS Applied Amount")
        rsWork("Discount") = rsAPDetail("AP CROSS Discount Taken")
        rsWork("Net Amt") = rsAPDetail("AP CROSS Applied Amount") + rsAPDetail("AP CROSS Discount Taken")
      rsWork.Update
      
      X% = X% + 1
      rsAPDetail.MoveNext
    Loop
  End If

  'On Error GoTo 0

  Dim BalancePaidForward#
  Dim BalanceDiscountForward#
  Dim BalanceNetForward#

  If X% > 10 Then 'Write a balance forward record
    BalancePaidForward# = SumRecord("[Amount Paid]", "[Print Check Work]", db, "[Visible] = FALSE")
    BalanceDiscountForward# = SumRecord("[Discount]", "[Print Check Work]", db, "[Visible] = False")
    BalanceNetForward# = BalancePaidForward# + BalanceDiscountForward#
    
    rsWork.AddNew
      rsWork("Vendor ID") = VendorID$
      rsWork("Check Number") = CheckNo$
      rsWork("Total Amount") = rsAPHeader("AP PAY Amount")
      rsWork("Transaction Date") = TDate
      rsWork("Order") = 0
      rsWork("Visible") = True
      rsWork("Reference #") = "Balance Forward"
      rsWork("Invoice #") = "N/A"
      rsWork("Invoice Date") = FormatDate(Now)
      rsWork("Invoice Amt") = 0
      rsWork("Amount Paid") = BalancePaidForward#
      rsWork("Discount") = BalanceDiscountForward#
      rsWork("Net Amt") = BalanceNetForward#
    rsWork.Update
  End If

  rsWork.Close

'
  'Add code to print checks here
'
  
  
  PrintCheck% = True


  Exit Function
  
PrintCheck_Error:
  Call ErrorLog("Purchase Module", "PrintCheck", Now, Err.Number, Err.Description, True, db)
  Resume Next

NoChecks:


End Function

Function RecurPayments(Optional db As ADODB.Connection) As Integer

'Dim Currentdb As Boolean
  
  'Dim db As ADODB.Connection
'  Currentdb = False
'  If db Is Nothing Then
'    Set db = New ADODB.Connection
'    db.CursorLocation = adUseServer
'    db.Open gblADOProvider
'    Currentdb = True
'  End If

  
'  Dim rsRecur As ADODB.Recordset
'  Set rsRecur = New ADODB.Recordset
'  rsRecur.Open "SELECT [AP PAY ID] FROM [AP Payment Header] where [AP PAY Recurring YN] = true", db, adOpenKeyset, adLockReadOnly, adCmdText

'  Dim DocumentKey&
'  Dim Success%
  'On Error Resume Next
  'rsRecur.MoveFirst
  'If Err <> 0 Then GoTo SkipRecurPayments
'  If rsRecur.RecordCount = 0 Then GoTo SkipRecurPayments
'  rsRecur.MoveFirst
  
'  Do While Not rsRecur.EOF
'    DocumentKey& = rsRecur("AP PAY ID")
'    Success% = ClonePayment(DocumentKey&, False, db)
'    rsRecur.MoveNext
'  Loop

'SkipRecurPayments:

'rsRecur.Close
'Set rsRecur = Nothing
'If Currentdb = True Then
'    db.Close
'    Set db = Nothing
'End If

End Function

Function RecurPurchases(Optional db As ADODB.Connection) As Integer
'Dim Currentdb As Boolean
  
  'Dim db As ADODB.Connection
'  Currentdb = False
'  If db Is Nothing Then
'    Set db = New ADODB.Connection
'    db.CursorLocation = adUseServer
'    db.Open gblADOProvider
'    Currentdb = True
'  End If

'  Dim rsRecur As ADODB.Recordset
'  Set rsRecur = New ADODB.Recordset
'  rsRecur.Open "SELECT [AP PO Document No] FROM [AP Purchase] where [AP PO Recurring YN] = true", db, adOpenKeyset, adLockReadOnly, adCmdText

'  Dim DocumentKey&
'  Dim Success%
  'On Error Resume Next
  'rsRecur.MoveFirst
  'If Err <> 0 Then GoTo SkipRecurPurchases
'  If rsRecur.RecordCount = 0 Then GoTo SkipRecurPurchases
'  rsRecur.MoveFirst
  
'  Do While Not rsRecur.EOF
'    DocumentKey& = rsRecur("AP PO Document No")
'    Success% = ClonePurchase(DocumentKey&, db)
'    rsRecur.MoveNext
'  Loop

'SkipRecurPurchases:

'rsRecur.Close
'Set rsRecur = Nothing
'If Currentdb = True Then
'    db.Close
'    Set db = Nothing
'End If

End Function


Public Function PostRMA(DocumentKey&, intShowError As Integer, db As ADODB.Connection)

'On Error GoTo PostRMA_Error
  'Dim db As ADODB.Connection
  'Set db = New ADODB.Connection
  'db.CursorLocation = adUseServer
  'db.Open gblADOProvider

  Dim msg$
  Dim title$

  Dim rsCompany As ADODB.Recordset
  Set rsCompany = New ADODB.Recordset
  rsCompany.Open "SELECT  [SYS COM Purchase AP Acct],[SYS COM Purchase Discount Acct]," & _
  "[SYS COM GL Post By Date],[SYS COM Purchase Inventory Acct],[SYS COM Purchase Freight Acct]," & _
  "[SYS COM Purchase Misc Acct] FROM [SYS Company]", db, adOpenKeyset, adLockOptimistic, admcdtext
  rsCompany.MoveFirst

  'Dim rsGLWorkDetail As ADODB.Recordset
  'Set rsGLWorkDetail = New ADODB.Recordset
  'rsGLWorkDetail.Open "[GL Work Detail]", db, adOpenStatic, adLockOptimistic, adCmdTable

  Dim rsPurchase As ADODB.Recordset
  Set rsPurchase = New ADODB.Recordset
  rsPurchase.Open "SELECT [AP PO Document Type],[AP PO Date],[AP PO Vendor ID],[AP PO Amount Paid]," & _
  "[AP PO Payment Method],[AP PO Check Number],[AP PO Check Acct ID],[AP PO Document No]," & _
  "[AP PO Ext Document No],[AP PO vendor Name],[AP PO Description],[AP PO Total Amount]," & _
  "[AP PO Shipping],[AP PO Misc Charges],[AP PO Vendor Invoice No],[AP PO Discount Amt] FROM [AP Purchase] " & _
  "WHERE [AP PO Document No]=" & DocumentKey&, db, adOpenKeyset, adLockOptimistic, adCmdText

  ' first lets get the po
  'rsPurchase.Index = "PrimaryKey"
  'rsPurchase.MoveFirst
  'rsPurchase.Find "[AP PO Document No]='" & DocumentKey& & "'"

  'Post by 1 - system date or 2 - Transaction date?
  Dim PostDate%
  PostDate% = rsCompany("SYS COM GL Post By Date")

  'Set Purchase Type
  Dim PurchaseType$
  PurchaseType$ = rsPurchase("AP PO Document Type")

  'Set Post Date
  Dim TranDate As Variant
  If PostDate% = 1 Then
    TranDate = DateValue(Format(Now, "Short Date"))
  Else
    TranDate = DateValue(rsPurchase("AP PO Date"))
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
    GoTo UnableToPostPOHere
  End If

  'On Error GoTo PostRMA_Error

  ' clear any GL Work records
  'Dim cmdtemp As ADODB.Recordset
  'Set cmdtemp = New ADODB.Recordset
  'cmdtemp.Open "DELETE DISTINCTROW * FROM [GL Work Detail]", db, , , adCmdText
  'cmdtemp.Close
  'Set cmdtemp = Nothing
  db.Execute "DELETE DISTINCTROW * FROM [GL Work Detail]"

  ' update vendor stats
  Dim rsVendor As ADODB.Recordset
  Set rsVendor = New ADODB.Recordset
  rsVendor.Open "SELECT [AP VEN ID], [AP VEN Name], [AP VEN Payments YTD]," & _
  "[AP VEN Payment Number Lifetime],[AP VEN Payment Number YTD],[AP VEN Purchase YTD], " & _
  "[AP VEN Purchase Lifetime],[AP VEN Purchase Number Lifetime],[AP VEN Purchase Number YTD]," & _
  "[AP VEN Payments Lifetime], [AP VEN Financial Period 1] FROM [AP Vendor] " & _
  "WHERE [AP VEN ID]='" & rsPurchase("AP PO Vendor ID") & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
  'rsVendor.Open "SELECT * FROM [AP Vendor] where [AP VEN ID] = '" & rsPurchase("AP PO Vendor ID") & "'", db, adOpenStatic, adLockOptimistic, adCmdText

  Dim CurrentBalance@
    If rsPurchase("AP PO Amount Paid") > 0 Then
      rsVendor("AP VEN Payments YTD") = rsVendor("AP VEN Payments YTD") - rsPurchase("AP PO Amount Paid")
      rsVendor("AP VEN Payments Lifetime") = rsVendor("AP VEN Payments Lifetime") - rsPurchase("AP PO Amount Paid")
      rsVendor("AP VEN Payment Number Lifetime") = rsVendor("AP VEN Payment Number Lifetime") - 1
      rsVendor("AP VEN Payment Number YTD") = rsVendor("AP VEN Payment Number YTD") - 1
    End If
    rsVendor("AP VEN Purchase YTD") = rsVendor("AP VEN Purchase YTD") - rsPurchase("AP PO Total Amount")
    rsVendor("AP VEN Purchase Lifetime") = rsVendor("AP VEN Purchase Lifetime") - rsPurchase("AP PO Total Amount")
    rsVendor("AP VEN Purchase Number Lifetime") = rsVendor("AP VEN Purchase Number Lifetime") - 1
    rsVendor("AP VEN Purchase Number YTD") = rsVendor("AP VEN Purchase Number YTD") - 1
    ' Update current Balance - if not Paid in full
    If rsPurchase("AP PO Total Amount") - rsPurchase("AP PO Amount Paid") > 0 Then
      CurrentBalance@ = IIf(IsNull(rsVendor("AP VEN Financial Period 1")), 0, rsVendor("AP VEN Financial Period 1"))
      CurrentBalance@ = CurrentBalance@ - (rsPurchase("AP PO Total Amount") + rsPurchase("AP PO Amount Paid"))
      rsVendor("AP VEN Financial Period 1") = CurrentBalance@
    End If
  rsVendor.Update
  rsVendor.Close
  Set rsVendor = Nothing

    Dim PaymentID&
    'Dim rsAPPaymentHeader As ADODB.Recordset
    'Set rsAPPaymentHeader = New ADODB.Recordset
    'rsAPPaymentHeader.Open "[AP Payment Header]", db, adOpenStatic, adLockOptimistic, adCmdTable
    
    'Dim rsAPCross As ADODB.Recordset
    'Set rsAPCross = New ADODB.Recordset
    'rsAPCross.Open "[AP Payment Invoice Cross Reference]", db, adOpenStatic, adLockOptimistic, adCmdTable
    
  If rsPurchase("AP PO Amount Paid") > 0 Then
    SQLstatement = "INSERT INTO [AP Payment Header]"
    SQLstatement = SQLstatement & " ([AP PAY Type],[AP PAY Check No],[AP PAY Vendor No]," & _
    "[AP PAY Transaction Date],[AP PAY Amount],[AP PAY UnApplied Amount],[AP PAY Credit YN]," & _
    "[AP PAY Bank Account],[AP PAY Status],[AP PAY Class],[AP PAY Posted YN],[AP PAY Void],[AP PAY Cleared])"
    
    '  write Payment Header
    SQLstatement = SQLstatement & " VALUES ('Credit Memo','" & _
    "RMA " & rsPurchase("AP PO Vendor Invoice No") & "','" & rsPurchase("AP PO Vendor ID") & "',#" & _
    rsPurchase("AP PO Date") & "#," & rsPurchase("AP PO Total Amount") & ","
    
    'rsAPPaymentHeader.AddNew
    '  rsAPPaymentHeader("AP PAY Type") = "Credit Memo"
    '  rsAPPaymentHeader("AP PAY Check No") = "RMA " & rsPurchase("AP PO Vendor Invoice No")
    '  rsAPPaymentHeader("AP PAY Vendor No") = rsPurchase("AP PO Vendor ID") & ""
    '  rsAPPaymentHeader("AP PAY Transaction Date") = rsPurchase("AP PO Date")
    '  rsAPPaymentHeader("AP PAY Amount") = rsPurchase("AP PO Total Amount")
  ' Disallow changes to RMA's that have been Refunded by check
      'If Forms![Purchase Transactions]![AP PO Amount Paid] > 0 Then
      If rsPurchase![AP PO Amount Paid] > 0 Then
        'rsAPPaymentHeader("AP PAY UnApplied Amount") = 0
        'rsAPPaymentHeader("AP PAY Credit YN") = False
        'rsAPPaymentHeader("AP PAY Bank Account") = rsPurchase("AP PO Check Acct ID")
        SQLstatement = SQLstatement & "0,False,'" & rsPurchase("AP PO Check Acct ID") & "',"
      Else
        'rsAPPaymentHeader("AP PAY UnApplied Amount") = rsPurchase("AP PO Total Amount")
        'rsAPPaymentHeader("AP PAY Credit YN") = True
        'rsAPPaymentHeader("AP PAY Bank Account") = "None"
        SQLstatement = SQLstatement & rsPurchase("AP PO Total Amount") & ",True,'None',"
      End If
      'rsAPPaymentHeader("AP PAY Status") = "Posted"
      'rsAPPaymentHeader("AP PAY Class") = 0 '
      'rsAPPaymentHeader("AP PAY Posted YN") = True
      'rsAPPaymentHeader("AP PAY Void") = False
      'rsAPPaymentHeader("AP PAY Cleared") = False
      SQLstatement = SQLstatement & "'Posted',0,True,False,False)"
      
      db.Execute SQLstatement
    'rsAPPaymentHeader.Update
    
    
      'PaymentID& = rsAPPaymentHeader("AP PAY ID")
    
    Else
      ' New AP Cross Payment and AP Payment Header
    
      '  write Payment Header
    SQLstatement = "INSERT INTO [AP Payment Header]"
    SQLstatement = SQLstatement & " ([AP PAY Type],[AP PAY Check No],[AP PAY Vendor No]," & _
    "[AP PAY Transaction Date],[AP PAY Amount],[AP PAY UnApplied Amount],[AP PAY Credit YN]," & _
    "[AP PAY Bank Account],[AP PAY Status],[AP PAY Class],[AP PAY Posted YN],[AP PAY Void],[AP PAY Cleared])"
      
      'rsAPPaymentHeader.AddNew
      '  rsAPPaymentHeader("AP PAY Type") = "Credit Memo"
      '  rsAPPaymentHeader("AP PAY Check No") = "RMA " & rsPurchase("AP PO Vendor Invoice No")
      '  rsAPPaymentHeader("AP PAY Vendor No") = rsPurchase("AP PO Vendor ID") & ""
      '  rsAPPaymentHeader("AP PAY Transaction Date") = rsPurchase("AP PO Date")
      '  rsAPPaymentHeader("AP PAY Amount") = rsPurchase("AP PO Total Amount")
      '  rsAPPaymentHeader("AP PAY UnApplied Amount") = rsPurchase("AP PO Total Amount")
      '  rsAPPaymentHeader("AP PAY Credit YN") = True
      '  rsAPPaymentHeader("AP PAY Bank Account") = "None"
      '  rsAPPaymentHeader("AP PAY Status") = "Posted"
      '  rsAPPaymentHeader("AP PAY Void") = False
      '  rsAPPaymentHeader("AP PAY Class") = 0 '"CreditMemo/RMA"
      '  rsAPPaymentHeader("AP PAY Cleared") = False
      '  rsAPPaymentHeader("AP PAY Posted YN") = True
      'rsAPPaymentHeader.Update
      ' end of write payment header
     SQLstatement = SQLstatement & " VALUES ('Credit Memo','" & _
     "RMA " & rsPurchase("AP PO Vendor Invoice No") & "','" & rsPurchase("AP PO Vendor ID") & "',#" & _
     rsPurchase("AP PO Date") & "#," & rsPurchase("AP PO Total Amount") & ","
     SQLstatement = SQLstatement & rsPurchase("AP PO Total Amount") & ",True,'None',"
     SQLstatement = SQLstatement & "'Posted',0,True,False,False)"
   
      ' End of New AP Cross Payment and AP Payment Header
  
  End If
      
      Dim rsAPPaymentHeader As ADODB.Recordset
      Set rsAPPaymentHeader = New ADODB.Recordset
      rsAPPaymentHeader.Open "SELECT [AP PAY ID] FROM [AP Payment Header] WHERE [AP PAY Check No]='RMA " & rsPurchase("AP PO Vendor Invoice No") & "'", db, adOpenKeyset, adLockReadOnly, adCmdText
      If rsAPPaymentHeader.RecordCount > 1 Then
        rsAPPaymentHeader.MoveLast
        PaymentID& = rsAPPaymentHeader("AP PAY ID")
      Else
        PaymentID& = rsAPPaymentHeader("AP PAY ID")
      End If
      rsAPPaymentHeader.Close
      Set rsAPPaymentHeader = Nothing

  ' Inventory Updates
  Dim rsInventory As ADODB.Recordset
  Dim rsDetail As ADODB.Recordset
  Set rsInventory = New ADODB.Recordset
  rsInventory.Open "SELECT [INV ITEM Id],[INV ITEM Last Cost],[INV ITEM Average Cost]," & _
  "[INV ITEM Qty On Hand],[INV ITEM Last Cost] FROM [INV Items]", db, adOpenKeyset, adLockOptimistic, adCmdTable
  
  Set rsDetail = New ADODB.Recordset
  rsDetail.Open "SELECT [AP POD Item ID],[AP POD Units],[AP POD Unit Cost]," & _
  "[AP POD Total Qty Received] FROM [AP Purchase Detail] where [AP POD Document No] = " & rsPurchase("AP PO Document No"), db, adOpenKeyset, adLockOptimistic, adCmdText

  'rsDetail.MoveLast
  'rsDetail.MoveFirst

  Dim Qty#
  Dim DetailUnitCost#
  Dim DetailItemCost#
  Dim CurrentLastCost#
  Dim CurrentAverageCost#
  Dim CurrentQuantityOnHand#
  Dim NewAverageCost#
  Dim Factor#
  Dim QtyRec#

  If rsDetail.RecordCount = 0 Then
    ' no detail data
  Else
    rsDetail.MoveFirst
    ' may have detail data
    Do While rsDetail.EOF = False
      'rsInventory.Index = "PrimaryKey"
      'rsInventory.Seek "=", rsDetail("AP POD Item ID")
      rsInventory.MoveFirst
      rsInventory.Find "[INV ITEM Id]='" & rsDetail("AP POD Item ID") & "'"
      'If rsInventory.BOF And rsInventory.EOF Then
      If rsInventory.EOF Then
        ' May be a non stock item
      Else
          ' Update Inventory Average & Last Cost
          Factor# = GetUOMMultiplier(rsDetail("AP POD Item ID"), rsDetail("AP POD Units"), db)
          
          DetailUnitCost# = IIf(IsNull(rsDetail("AP POD Unit Cost")), 0, rsDetail("AP POD Unit Cost"))
          If Factor# = 0 Then
          Else
            DetailItemCost# = DetailUnitCost# / Factor#
          End If
          'update costs on returns
          If DetailUnitCost# > 0 Then
            CurrentLastCost# = IIf(IsNull(rsInventory("INV ITEM Last Cost")), 0, rsInventory("INV ITEM Last Cost"))
            CurrentAverageCost# = IIf(IsNull(rsInventory("INV ITEM Average Cost")), 0, rsInventory("INV ITEM Average Cost"))
            CurrentQuantityOnHand# = IIf(IsNull(rsInventory("INV ITEM Qty On Hand")), 0, rsInventory("INV ITEM Qty On Hand"))
            rsInventory("INV ITEM Last Cost") = DetailItemCost#
            
            QtyRec# = IIf(IsNull(rsDetail("AP POD Total Qty Received")), 0, rsDetail("AP POD Total Qty Received"))
            QtyRec# = QtyRec# * Factor#
            If CurrentQuantityOnHand# < 0 Then CurrentQuantityOnHand# = 0
            NewAverageCost# = ((CurrentQuantityOnHand# * CurrentAverageCost#) - (QtyRec# * DetailItemCost#))
            If CurrentQuantityOnHand# + QtyRec# <= 0 Then
              NewAverageCost# = DetailItemCost#
            Else
              NewAverageCost# = NewAverageCost# / (CurrentQuantityOnHand# - QtyRec#)
            End If
            rsInventory("INV ITEM Average Cost") = NewAverageCost#
          End If
          ' end of Update Inventory Average & Last Cost
          
          'Take units into consideration
          rsInventory("INV ITEM Qty On Hand") = rsInventory("INV ITEM Qty On Hand") - rsDetail("AP POD Total Qty Received") * Factor#
        rsInventory.Update
      End If
  
      rsDetail.MoveNext
    Loop
    rsInventory.Close
    Set rsInventory = Nothing
    
    rsDetail.Close
    Set rsDetail = Nothing
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
    
    'rsGLTrans("GL TRANS Document #") = "RMA " & rsPurchase("AP PO Ext Document No")
    
    ' gl post date
    Dim TempStr As String
    If PostDate% = 1 Then
      TempStr = Format(Now, "Short Date")
    Else
      TempStr = rsPurchase("AP PO Date")
    End If
    
    'rsGLTrans("GL TRANS Type") = "RMA"
    
    SQLstatement = SQLstatement & " VALUES ('RMA " & rsPurchase("AP PO Ext Document No") & "','RMA',#" & TempStr & "#,"
    
    refr$ = rsPurchase("AP PO vendor Name")
    desc$ = IIf(IsNull(rsPurchase("AP PO Description")), "", rsPurchase("AP PO Description"))
    If Len(Trim$(desc$)) = 0 Then
      desc$ = "RMA " & rsPurchase("AP PO Ext Document No")
    End If
    
      SQLstatement = SQLstatement & "'" & refr$ & "'," & rsPurchase("AP PO Total Amount") & ",1,"
      SQLstatement = SQLstatement & "'" & desc$ & "','RMA " & rsPurchase("AP PO Ext Document No") & "',True)"
      'Debug.Print SQLstatement
      
      db.Execute SQLstatement
    
  '  rsGLTrans("GL TRANS Reference") = refr$
  '  rsGLTrans("GL TRANS Amount") = rsPurchase("AP PO Total Amount")
  '  rsGLTrans("GL TRANS Posted YN") = 1
  '  rsGLTrans("GL TRANS Description") = desc$
  '  rsGLTrans("GL TRANS Source") = "RMA " & rsPurchase("AP PO Ext Document No")
  '  rsGLTrans("GL TRANS System Generated") = True
  'rsGLTrans.Update
  Dim rsGLTrans As ADODB.Recordset
  Set rsGLTrans = New ADODB.Recordset
  rsGLTrans.Open "SELECT [GL TRANS Number] FROM [GL Transaction] WHERE [GL TRANS Document #]='" & "RMA " & rsPurchase("AP PO Ext Document No") & "'", db, adOpenKeyset, adLockReadOnly, adCmdText
      NewNumber& = rsGLTrans("GL TRANS Number")
  rsGLTrans.Close
  Set rsGLTrans = Nothing

  ' update GL
  '-----------------------------------------------------------------------
  ' PO with Payment GL Affected Accounts
  '
  '                  Debit   Credit   Source
  '                  -----   ------   ------
  ' Inventory                  x       Item - Inventory
  ' Misc Charges               x       Pref - Purchases
  ' Freight Expense            x       Pref - Purchases
  ' Discount           x               Pref - Purchases
  ' AP                 x               Pref - Purchases
  ' CASH               x               Bank - Cash Acct
  ' The following entry is only valid if the payment exceeds the receipt total
  ' AP                         x       Pref - Purchases
  '
  ' Notes:
  ' Each Inventory Item is processed and the GL Acct is Retrieved.
  '-----------------------------------------------------------------------

  Dim InventoryAcct$

  ' Debits
  ' Inventory Increase
  Dim Longer&
  Longer& = 0
  Set rsDetail = New ADODB.Recordset
  'rsDetail.Open "SELECT * FROM [AP Purchase Detail] where [AP POD Document No] = " & rsPurchase("AP PO Document No"), db, adOpenStatic, adLockOptimistic, adCmdText
  rsDetail.Open "SELECT [AP POD Posting Account],[AP POD Item ID],[AP POD Item Total] " & _
  "FROM [AP Purchase Detail] where [AP POD Document No] = " & rsPurchase("AP PO Document No"), db, adOpenKeyset, adLockOptimistic, adCmdText
  'rsDetail.MoveLast
  'rsDetail.MoveFirst

  If rsDetail.RecordCount = 0 Then
    ' no detail data
  Else
    ' may have detail data
    rsDetail.MoveFirst
    Do While rsDetail.EOF = False
        
        '+++
        ' use Account in detail 1st
        InventoryAcct$ = IIf(IsNull(rsDetail("AP POD Posting Account")), "", rsDetail("AP POD Posting Account"))

        ' use Account in Item 2nd
        If Len(InventoryAcct$) = 0 Then
          'rsInventory.Index = "PrimaryKey"
          'rsInventory.Seek "=", rsDetail("AP POD Item ID")
          Set rsInventory = New ADODB.Recordset
          rsInventory.Open "SELECT [INV ITEM Inventory Account] FROM FROM [INV Items] WHERE [INV ITEM Id] ='" & rsDetail("AP POD Item ID") & "'", db, adOpenKeyset, adLockReadOnly, adCmdText
          'If rsInventory.BOF And rsInventory.EOF Then
          If rsInventory.RecordCount = 0 Then
          Else
            InventoryAcct$ = IIf(IsNull(rsInventory("INV ITEM Inventory Account")), "", rsInventory("INV ITEM Inventory Account"))
          End If
          rsInventory.Close
          Set rsInventory = Nothing
        End If

        ' use Account in Vendor 3rd
        If Len(InventoryAcct$) = 0 Then
          Dim VendorKey$
          
          VendorKey$ = IIf(IsNull(rsPurchase("AP PO Vendor ID")), "", rsPurchase("AP PO Vendor ID"))
          If Len(VendorKey$) > 0 Then
            Set rsVendor = New ADODB.Recordset
            rsVendor.Open "SELECT [AP VEN Default GL] FROM [AP Vendor] WHERE [AP VEN ID]='" & VendorKey$ & "'", db, adOpenKeyset, adLockReadOnly, adCmdText
            'Set rsVendor = db2.OpenRecordset("AP Vendor")
            'rsVendor.Index = "PrimaryKey"
            'rsVendor.Seek "=", VendorKey$
            If rsVendor.RecordCount > 0 Then
              InventoryAcct$ = IIf(IsNull(rsVendor("AP VEN Default GL")), "", rsVendor("AP VEN Default GL"))
            Else
              InventoryAcct$ = ""
            End If
            rsVendor.Close
            Set rsVendor = Nothing
          End If
        End If

        'Use preferences Acct 4th
        If Len(InventoryAcct$) = 0 Then InventoryAcct$ = rsCompany("SYS COM Purchase Inventory Acct")
        '+++
      SQLstatement = "INSERT INTO [GL Work Detail]"
      SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
      SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & InventoryAcct$ & "',0," & rsDetail("AP POD Item Total") & ")"
      db.Execute SQLstatement
        
        'rsGLWorkDetail.AddNew
        '  rsGLWorkDetail("GW TRANSD Number") = NewNumber&
        '  rsGLWorkDetail("GW TRANSD Account") = InventoryAcct$
        '  rsGLWorkDetail("GW TRANSD Debit Amount") = 0
        '  rsGLWorkDetail("GW TRANSD Credit Amount") = rsDetail("AP POD Item Total")
        '  rsGLWorkDetail("GW TRANSD Project") = ""
        'rsGLWorkDetail.Update

      rsDetail.MoveNext
    Loop
  End If
  rsDetail.Close
  Set rsDetail = Nothing
  
  ' Freight Expense
  If rsPurchase("AP PO Shipping") > 0 Then
      SQLstatement = "INSERT INTO [GL Work Detail]"
      SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
      SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & rsCompany("SYS COM Purchase Freight Acct") & "',0," & rsPurchase("AP PO Shipping") & ")"
      db.Execute SQLstatement
    'rsGLWorkDetail.AddNew
    '  rsGLWorkDetail("GW TRANSD Number") = NewNumber&
    '  rsGLWorkDetail("GW TRANSD Account") = rsCompany("SYS COM Purchase Freight Acct")
    '  rsGLWorkDetail("GW TRANSD Debit Amount") = 0
    '  rsGLWorkDetail("GW TRANSD Credit Amount") = rsPurchase("AP PO Shipping")
    '  rsGLWorkDetail("GW TRANSD Project") = ""
    'rsGLWorkDetail.Update
  End If

  ' Misc Charges (Post Debit for Restocking Charges)
  If rsPurchase("AP PO Misc Charges") < 0 Then
      SQLstatement = "INSERT INTO [GL Work Detail]"
      SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
      SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & rsCompany("SYS COM Purchase Misc Acct") & "'," & rsPurchase("AP PO Misc Charges") * -1 & ",0)"
      db.Execute SQLstatement
    'rsGLWorkDetail.AddNew
    '  rsGLWorkDetail("GW TRANSD Number") = NewNumber&
    '  rsGLWorkDetail("GW TRANSD Account") = rsCompany("SYS COM Purchase Misc Acct")
    '  rsGLWorkDetail("GW TRANSD Debit Amount") = rsPurchase("AP PO Misc Charges") * -1
    '  rsGLWorkDetail("GW TRANSD Credit Amount") = 0
    '  rsGLWorkDetail("GW TRANSD Project") = ""
    'rsGLWorkDetail.Update
  End If
  
  ' Credits
  ' Discount Amount
  If rsPurchase("AP PO Discount Amt") > 0 Then
      SQLstatement = "INSERT INTO [GL Work Detail]"
      SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
      SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & rsCompany("SYS COM Purchase Discount Acct") & "'," & rsPurchase("AP PO Discount Amt") & ",0)"
      db.Execute SQLstatement
    'rsGLWorkDetail.AddNew
    '  rsGLWorkDetail("GW TRANSD Number") = NewNumber&
    '  rsGLWorkDetail("GW TRANSD Account") = rsCompany("SYS COM Purchase Discount Acct")
    '  rsGLWorkDetail("GW TRANSD Debit Amount") = rsPurchase("AP PO Discount Amt")
    '  rsGLWorkDetail("GW TRANSD Credit Amount") = 0
    '  rsGLWorkDetail("GW TRANSD Project") = ""
    'rsGLWorkDetail.Update
  End If
  
  ' Cash Payment
  If rsPurchase("AP PO Amount Paid") > 0 Then
      SQLstatement = "INSERT INTO [GL Work Detail]"
      SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
      SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & rsPurchase("AP PO Check Acct ID") & "'," & rsPurchase("AP PO Amount Paid") & ",0)"
      db.Execute SQLstatement
    'rsGLWorkDetail.AddNew
    '  rsGLWorkDetail("GW TRANSD Number") = NewNumber&
    '  rsGLWorkDetail("GW TRANSD Account") = rsPurchase("AP PO Check Acct ID")
    '  rsGLWorkDetail("GW TRANSD Debit Amount") = rsPurchase("AP PO Amount Paid")
    '  rsGLWorkDetail("GW TRANSD Credit Amount") = 0
    '  rsGLWorkDetail("GW TRANSD Project") = ""
    'rsGLWorkDetail.Update

  End If

  ' AP
  If rsPurchase("AP PO Total Amount") - rsPurchase("AP PO Amount Paid") > 0 Then
      SQLstatement = "INSERT INTO [GL Work Detail]"
      SQLstatement = SQLstatement & " ([GW TRANSD Number],[GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount])"
      SQLstatement = SQLstatement & " VALUES (" & NewNumber& & ",'" & rsCompany("SYS COM Purchase AP Acct") & "'," & rsPurchase("AP PO Total Amount") - rsPurchase("AP PO Amount Paid") & ",0)"
      db.Execute SQLstatement
    'rsGLWorkDetail.AddNew
    '  rsGLWorkDetail("GW TRANSD Number") = NewNumber&
    '  rsGLWorkDetail("GW TRANSD Account") = rsCompany("SYS COM Purchase AP Acct")
    '  rsGLWorkDetail("GW TRANSD Debit Amount") = rsPurchase("AP PO Total Amount") - rsPurchase("AP PO Amount Paid")
    '  rsGLWorkDetail("GW TRANSD Credit Amount") = 0
    '  rsGLWorkDetail("GW TRANSD Project") = ""
    'rsGLWorkDetail.Update
  End If


SkipIt2:


PostRMA_Exit:

  ' post GL entry this receiving
  Dim Success%
  Success% = PostGLWorkDetail(TranDate, NewNumber&, db)
  If Success% = False Then
    MsgBox "An error occurred writing GL Transaction!", , "Error"
    PostRMA = False
    Exit Function
  End If

  PostRMA = True
  
   rsCompany.Close
   Set rsCompany = Nothing
   rsPurchase.Close
   Set rsPurchase = Nothing

  Exit Function

PostRMA_Error:
  Call ErrorLog("Purchase Module", "PostRMA", Now, Err.Number, Err.Description, intShowError, db)
  PostRMA = False
  Resume Next

UnableToPostPOHere:
  PostRMA = False
   rsCompany.Close
   Set rsCompany = Nothing
   rsPurchase.Close
   Set rsPurchase = Nothing
  Exit Function

End Function

Public Function BuildAPAging()
  Dim db As ADODB.Connection
  Set db = New ADODB.Connection
  db.CursorLocation = adUseClient
  db.Open gblADOProvider

Dim rstInfo As ADODB.Recordset

Set rstInfo = New ADODB.Recordset
rstInfo.Open "[exp qryGetPeriod-n-AgingInfo]", db, adOpenStatic, adLockOptimistic

intPeriod1 = rstInfo("SYS COM Purchase Period 1")
intPeriod2 = rstInfo("SYS COM Purchase Period 2")
intPeriod3 = rstInfo("SYS COM Purchase Period 3")
DoCmd.SetWarnings False

db.Execute ("Delete * From [Print Aged Payables Work]")

'Find Proper Period for Amounts
'Pull Discounts Applied to Aged Purchases
DoCmd.OpenQuery "exp qry Process APR Discounts"
'Pull Payments on Aged Purchases
DoCmd.OpenQuery "exp qry Process APR Payments"
'Pull Unapplied Credits/RMA's
DoCmd.OpenQuery "exp - qryAppendCreditRMAtoWorkTable"
'Pull Unapplied Payments
DoCmd.OpenQuery "exp - qryAppendUnappliedAPtoWorkTable"
'Now add the actual Purchases
DoCmd.OpenQuery "exp - qryAddAgedPurchasestoWorkTable"
DoCmd.SetWarnings True
'Show the Report
DoCmd.OpenReport "AP - Aged Payables Detail", A_PREVIEW
End Function

Public Sub GetVendorFinancials(VendorID$, db As ADODB.Connection, VendorP1Balance As Currency, VendorP2Balance As Currency, VendorP3Balance As Currency, VendorP4Balance As Currency, VendorPtotalBalance As Currency)

  'On Error GoTo GetVendorFinancials_Error

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

  'On Error GoTo VendorFinancial_Error

  VendorP1Balance = 0
  VendorP2Balance = 0
  VendorP3Balance = 0
  VendorP4Balance = 0
  VendorPtotalBalance = 0

  Dim rsCompany As ADODB.Recordset
  Set rsCompany = New ADODB.Recordset
  rsCompany.Open "SELECT [SYS COM Purchase Period 1],[SYS COM Purchase Period 2]," & _
  "[SYS COM Purchase Period 3],[SYS COM Purchase Age Invoices By] " & _
  "FROM [SYS Company]", db, adOpenKeyset, adLockOptimistic, adCmdText

  rsCompany.MoveFirst
  Period1% = rsCompany("SYS COM Purchase Period 1")
  Period2% = rsCompany("SYS COM Purchase Period 2")
  Period3% = rsCompany("SYS COM Purchase Period 3")
  Period4% = 90
  AgeBy% = IIf(IsNull(rsCompany("SYS COM Purchase Age Invoices By")), 1, rsCompany("SYS COM Purchase Age Invoices By"))
  '1 - Invoice Date  2 - Due Date
  
  rsCompany.Close
  Set rsCompany = Nothing
  
  'Go through AP Purchase and get transactions for this Vendor
  '   with balances > 0
  
  Dim dn As Recordset
  Dim rsAPPay As ADODB.Recordset
  Set dn = New ADODB.Recordset
  dn.Open "SELECT [AP PO Ext Document No],[AP PO Document Type],[AP PO Date],[AP PO Due Date]," & _
  "[AP PO Vendor ID],[AP PO Vendor Name],[AP PO Total Amount],[AP PO Balance Due] FROM [AP Purchase] " & _
  "WHERE [AP PO Vendor ID] = '" & VendorID$ & "' AND [AP PO Balance Due] > 0 AND [AP PO Posted YN] = TRUE", db, adOpenKeyset, adLockOptimistic, adCmdText
  
  db.Execute "DELETE * FROM [AGE Aging Purchase Work] WHERE [AGE Vendor ID]='" & VendorID$ & "'", , admcdtext
  
  'On Error Resume Next
  If dn.RecordCount > 0 Then
    dn.MoveFirst
    Do While Not dn.EOF
      'Get the balance
      Balance# = IIf(IsNull(dn("AP PO Balance Due")), 0, dn("AP PO Balance Due"))

      'Get Transaction Type to see if we sould increase or decrease the balance
      TransType$ = dn("AP PO Document Type")
      Select Case TransType$
      Case "Receiving", "Voucher", "Beginning Balance"
      Case "Credit Memo"
        'Find payment information and back out unapplied amount
        'rsAPPay.Index = "PrimaryKey"
        Set rsAPPay = New ADODB.Recordset
        rsAPPay.Open "SELECT [AP PAY Unapplied Amount] FROM [AP Payment Header] " & _
        "WHERE [AP PAY Check No]='CM " & dn("AP PO Ext Document No") & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
        'rsAPPay.Seek "=", dn("AP PO Vendor ID"), dn("AP PO Ext Document No")
        
        If rsAPPay.RecordCount > 0 Then
          Balance# = rsAPPay("AP PAY Unapplied Amount") * -1
        End If
        
        rsAPPay.Close
        Set rsAPPay = Nothing
        
      Case Else
        Balance# = 0
      End Select

      'Get a date to age by
      If (AgeBy% = 1) Then 'Use Invoice Date
        TransDate = IIf(IsNull(dn("AP PO Date")), FormatDate(FormatDate("1/01/" & Format(Now, "yyyy")), "Short Date"), dn("AP PO Date"))
      Else                 'Use Due Date
        TransDate = IIf(IsNull(dn("AP PO Due Date")), FormatDate(FormatDate("1/01/" & Format(Now, "yyyy")), "Short Date"), dn("AP PO Due Date"))
      End If

      Days& = DateDiff("d", TransDate, Now)
      
      SQLstatement = "INSERT INTO [AGE Aging Purchase Work]"
      SQLstatement = SQLstatement & " ([AGE Vendor ID],[AGE Vendor Name]," & _
      "[AGE PO Doc Ext No],[AGE Start Date],[AGE Orig Amount],[AGE Period 1]," & _
      "[AGE Period 2],[AGE Period 3],[AGE Period 4])"
        
      SQLstatement = SQLstatement & " VALUES ('" & dn![AP PO Vendor ID] & "','" & _
      dn![AP PO Vendor Name] & "','" & dn![AP PO Ext Document No] & "',#" & _
      TransDate & "#," & dn![AP PO Total Amount] & ","
      
      Select Case Days&
      Case Is < 0
        'Add it as current
        SQLstatement = SQLstatement & dn![AP PO Balance Due] & ",0,0,0" & ")"
        db.Execute SQLstatement, , adCmdText
        VendorP1Balance = VendorP1Balance + Balance#
      Case 0 To Period1%
        SQLstatement = SQLstatement & dn![AP PO Balance Due] & ",0,0,0" & ")"
        db.Execute SQLstatement, , adCmdText
        VendorP1Balance = VendorP1Balance + Balance#
      Case Period1% To Period2%
        SQLstatement = SQLstatement & "0," & dn![AP PO Balance Due] & ",0,0" & ")"
        db.Execute SQLstatement, , adCmdText
        VendorP2Balance = VendorP2Balance + Balance#
      Case Period2% To Period3%
        SQLstatement = SQLstatement & "0,0," & dn![AP PO Balance Due] & ",0" & ")"
        db.Execute SQLstatement, , adCmdText
        VendorP3Balance = VendorP3Balance + Balance#
      Case Else
        SQLstatement = SQLstatement & "0,0,0," & dn![AP PO Balance Due] & ")"
        db.Execute SQLstatement, , adCmdText
        VendorP4Balance = VendorP4Balance + Balance#
      End Select
      dn.MoveNext
    Loop
  End If
  
  dn.Close
  Set dn = Nothing
  
  'Now do payments
  Set dn = New ADODB.Recordset
  dn.Open "SELECT [AP PAY UnApplied Amount],[AP PAY Transaction Date] FROM [AP PAYMENT Header] where [AP PAY Vendor No] = '" & VendorID$ & "' AND [AP PAY Void] = False AND [AP PAY Posted YN] = TRUE", db, adOpenKeyset, adLockOptimistic, adCmdText
  
  'On Error Resume Next
  If dn.RecordCount > 0 Then
    dn.MoveFirst
    Do While Not dn.EOF
      'Get the balance
      Balance# = IIf(IsNull(dn("AP PAY UnApplied Amount")), 0, dn("AP PAY UnApplied Amount"))
      'Back out payment if type is return or NSF
      Select Case "AP PAY Type"
      Case "Void"
        Balance# = Balance# * -1
      Case "Credit Memo"
        Balance# = 0
      End Select

      TransDate = IIf(IsNull(dn("AP PAY Transaction Date")), FormatDate(FormatDate("1/01/" & Format(Now, "yyyy")), "Short Date"), dn("AP PAY Transaction Date"))
      
      Days& = DateDiff("d", TransDate, Now)

      Select Case Days&
      Case Is < 0
        'Don't use it
      Case 0 To Period1%
        VendorP1Balance = VendorP1Balance - Balance#
      Case Period1% To Period2%
        VendorP2Balance = VendorP2Balance - Balance#
      Case Period2% To Period3%
        VendorP3Balance = VendorP3Balance - Balance#
      Case Else
        VendorP4Balance = VendorP4Balance - Balance#
      End Select
      dn.MoveNext
    Loop
  End If
  dn.Close
  Set dn = Nothing
  
  VendorPtotalBalance = VendorP1Balance + VendorP2Balance + VendorP3Balance + VendorP4Balance

VendorFinancials_Exit:
  Exit Sub

VendorFinancial_Error:
  Dim msg$
  msg$ = Error(Err)
  MsgBox msg$, , "Vendor Financials"
  Exit Sub

GetVendorFinancials_Error:
  Call ErrorLog("Vendor Financial Tab", "GetVendorFinancials", Now, Err.Number, Err.Description, True, db)
  Resume Next

End Sub


