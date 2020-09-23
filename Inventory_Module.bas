Attribute VB_Name = "Inventory_Module"

Sub CalculateOnOrder()

  'On Error GoTo CalculateOnOrder_Error

    
  Dim db As ADODB.Connection
  Set db = New ADODB.Connection
  db.CursorLocation = adUseClient
  db.Open gblADOProvider
  
  Dim rsItems As ADODB.Recordset
  rsItems.Open "INV Items", db, adOpenStatic, adLockOptimistic, adCmdTable
    If rsItems.BOF = True And rsItems.EOF = True Then Exit Sub
  
  Dim dblQtyOnOrder As Double
  
  rsItems.MoveFirst
  Do While Not rsItems.EOF
    dblQtyOnOrder = NZ(SumRecord("[AR ORDERD Qty]", "[AR Order Detail]", db, "[AR ORDERD Item ID] = '" & rsItems("INV ITEM ID") & "'"))
      rsItems("INV ITEM Qty On Order") = dblQtyOnOrder
    rsItems.Update
    rsItems.MoveNext
  Loop
  
  rsItems.Close
  Set rsItems = Nothing
  db.Close
  Set db = Nothing

  Exit Sub
CalculateOnOrder_Error:
  Call ErrorLog("Inventory Module", "CalculateOnOrder", Now, Err.Number, Err.Description, True, db)
  Resume Next
    
  rsItems.Close
  Set rsItems = Nothing
  db.Close
  Set db = Nothing

End Sub


Function PostAdjustment(AdjustmentID&, Optional db As ADODB.Connection)
  
  'On Error GoTo PostAdjustment_Error
Dim Currentdb As Boolean
  
  'Dim db As ADODB.Connection
  Currentdb = False
  If db Is Nothing Then
    Set db = New ADODB.Connection
    db.CursorLocation = adUseServer
    db.Open gblADOProvider
    Currentdb = True
  End If
  
  Dim msg$
  Dim title$
  
  Dim rsCompany As ADODB.Recordset
  Set rsCompany = New ADODB.Recordset
  rsCompany.Open "SELECT  [SYS COM GL Post By Date],[SYS COM Sales COGS Acct] " & _
  "FROM [SYS Company]", db, adOpenKeyset, adLockOptimistic, adCmdText
  rsCompany.MoveFirst
'  rsCompany.Open "[SYS Company]", db, adOpenStatic, adLockOptimistic, adCmdTable
'  rsCompany.MoveFirst

  Dim rsAdjustment As ADODB.Recordset
  Set rsAdjustment = New ADODB.Recordset
  rsAdjustment.Open "SELECT [INV ADJ Document No],[INV ADJ Type],[INV ADJ Post To GL YN]," & _
  "[INV ADJ Ext Document No],[INV ADJ Date],[INV ADJ Reason] FROM [INV Adjustment] " & _
  "WHERE [INV ADJ Document No]=" & AdjustmentID&, db, adOpenKeyset, adLockOptimistic, adCmdText
  
  ' first lets get the adjustment
  'rsAdjustment.Index = "PrimaryKey"
  'rsAdjustment.Seek AdjustmentID&
  'rsAdjustment.Find "[INV ADJ Document No]=" & AdjustmentID&

  'Post by 1 - system date or 2 - Transaction date?
  Dim PostDate%
  PostDate% = rsCompany("SYS COM GL Post By Date")

  'Set Adjustment Type
  Dim AdjustmentType$
  AdjustmentType$ = rsAdjustment("INV ADJ Type")

  'Set Post Date
  Dim TranDate As Variant
  If PostDate% = 1 Then
    TranDate = DateValue(Format(Now, "Short Date"))
  Else
    TranDate = DateValue(rsAdjustment("INV ADJ Date"))
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
    GoTo UnableToPostAdjustmentHere
  End If

  'On Error GoTo PostAdjustment_Error

  ' clear any GL Work records
  'Dim cmdtemp As ADODB.Recordset
  'Set cmdtemp = New ADODB.Recordset
  'cmdtemp.Open "DELETE DISTINCTROW * FROM [GL Work Detail]", db, , , adCmdText
  'cmdtemp.Close
  'Set cmdtemp = Nothing
  
  db.Execute "DELETE DISTINCTROW * FROM [GL Work Detail]", , adCmdText
  
  'Write GL Header Record
  'Dim rsGLTrans As ADODB.Recordset
  'Set rsGLTrans = New ADODB.Recordset
  'rsGLTrans.Open "[GL Transaction]", db, adOpenStatic, adLockOptimistic, adCmdTable

  'Dim rsGLWorkDetail As ADODB.Recordset
  'Set rsGLWorkDetail = New ADODB.Recordset
  'rsGLWorkDetail.Open "[GL Work Detail]", db, adOpenStatic, adLockOptimistic, adCmdTable

  Dim NewNumber&

  If rsAdjustment("INV ADJ Post To GL YN") = True Then
    'rsGLTrans.AddNew
    '  rsGLTrans("GL TRANS Document #") = "INV ADJ " & rsAdjustment("INV ADJ Ext Document No")
    '  rsGLTrans("GL Trans Type") = "Inv Adjustment"
      
      Dim SQLstatement As String
      
      SQLstatement = "INSERT INTO [GL Transaction]"
      SQLstatement = SQLstatement & " ([GL TRANS Document #],[GL TRANS Type],[GL TRANS Date],"
      SQLstatement = SQLstatement & " [GL TRANS Description],[GL TRANS Amount],[GL TRANS Reference],"
      SQLstatement = SQLstatement & " [GL TRANS Source],[GL TRANS Posted YN],[GL TRANS System Generated])"
    
      Dim TempStr As String
    
      ' gl post date
      If PostDate% = 1 Then
        TempStr = Format(Now, "Short Date")
      Else
        TempStr = rsAdjustment("INV ADJ Date")
      End If
      
      SQLstatement = SQLstatement & " VALUES ('INV ADJ " & rsAdjustment("INV ADJ Ext Document No") & "','Inv Adjustment',#" & TempStr & "#,"
      SQLstatement = SQLstatement & "'INVADJ " & rsAdjustment("INV ADJ Ext Document No") & "'," & SumRecord("[INV ADJD Cost]", "[INV Adjustment Detail]", db, "[INV ADJD Document No] = " & AdjustmentID&) & ",'" & rsAdjustment("INV ADJ Reason") & "',"
      SQLstatement = SQLstatement & "'INVADJ " & rsAdjustment("INV ADJ Ext Document No") & "',1,True)"
      'Debug.Print SQLstatement
      
      db.Execute SQLstatement
  
      'rsGLTrans("GL Trans Description") = "INVADJ " & rsAdjustment("INV ADJ Ext Document No")
      'rsGLTrans("GL Trans Amount") = SumRecord("[INV ADJD Cost]", "[INV Adjustment Detail]", "[INV ADJD Document No] = " & AdjustmentID&)
      'rsGLTrans("GL Trans Reference") = rsAdjustment("INV ADJ Reason")
      'rsGLTrans("GL Trans Source") = "INVADJ " & rsAdjustment("INV ADJ Ext Document No")
      'rsGLTrans("GL Trans Posted YN") = 1
      'rsGLTrans("GL TRANS System Generated") = True
    'rsGLTrans.Update
    Dim rsGLTrans As ADODB.Recordset
    Set rsGLTrans = New ADODB.Recordset
    rsGLTrans.Open "SELECT [GL TRANS Number] FROM [GL Transaction]WHERE [GL TRANS Document #]='" & "INV ADJ " & rsAdjustment("INV ADJ Ext Document No") & "'", db, adOpenKeyset, adLockReadOnly, adCmdText
        NewNumber& = rsGLTrans("GL TRANS Number")
    rsGLTrans.Close
    Set rsGLTrans = Nothing
  End If

  Dim mTotalCost@
  mTotalCost@ = 0

  Dim rsInventory As ADODB.Recordset
  'Set rsInventory = New ADODB.Recordset
  'rsInventory.Open "SELECT [INV ITEM Id],[INV ITEM Qty On Hand]," & _
  "[INV ITEM Cost of Sales Account],[INV ITEM Average Cost],[INV ITEM Last Cost] " & _
  "FROM [INV Items]", db, adOpenKeyset, adLockOptimistic, adCmdText
  'rsInventory.Index = "PrimaryKey"

  Dim rsAdjustmentDetail As ADODB.Recordset
  Set rsAdjustmentDetail = New ADODB.Recordset
  rsAdjustmentDetail.Open "SELECT [INV ADJD Item ID],[INV ADJD Cost]," & _
  "[INV ADJD Adjusted Qty],[INV ADJD Posting Account] FROM [INV Adjustment Detail] " & _
  "WHERE [INV ADJD Document No] = " & AdjustmentID&, db, adOpenKeyset, adLockOptimistic, adCmdText

  rsAdjustmentDetail.MoveFirst

  Dim TCost#
  Dim AdjQty!
  Dim OrgQty!
  Dim CostAcct$
  Dim TOrgQty!
  Dim AvgCost#
  Dim NewQty#

  Do While Not rsAdjustmentDetail.EOF
    If Len(Trim(rsAdjustmentDetail("INV ADJD Item ID"))) = 0 Then
    Else
      TCost# = rsAdjustmentDetail("INV ADJD Cost")
      AdjQty! = rsAdjustmentDetail("INV ADJD Adjusted Qty") 'CSng(AdjustedQty(x%))
      
      Set rsInventory = New ADODB.Recordset
      rsInventory.Open "SELECT [INV ITEM Id],[INV ITEM Qty On Hand]," & _
      "[INV ITEM Cost of Sales Account],[INV ITEM Average Cost],[INV ITEM Last Cost] " & _
      "FROM [INV Items] WHERE [INV ITEM Id]='" & rsAdjustmentDetail("INV ADJD Item ID") & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
      rsInventory.MoveFirst
      'rsInventory.Find "[INV ITEM Id]='" & rsAdjustmentDetail("INV ADJD Item ID") & "'"

      If rsInventory.RecordCount = 0 Then
        'This is a problem
      Else
        OrgQty! = CSng(IIf(IsNull(rsInventory("INV ITEM Qty On Hand")), 0, rsInventory("INV ITEM Qty On Hand")))
        CostAcct$ = IIf(IsNull(rsInventory("INV ITEM Cost of Sales Account")), "", rsInventory("INV ITEM Cost of Sales Account"))
        If CostAcct$ = "" Then
          CostAcct$ = rsCompany("SYS COM Sales COGS Acct")
        End If
        If OrgQty! < 0 Then
          TOrgQty! = 0
        Else
          TOrgQty! = OrgQty!
        End If
        If (AdjustmentType$ = "Increase") Then
          AvgCost# = IIf(IsNull(rsInventory("INV ITEM Average Cost")), 0, rsInventory("INV ITEM Average Cost"))
          AvgCost# = (AvgCost# * TOrgQty!) + (TCost# * AdjQty!)
          If OrgQty! + AdjQty! <= 0 Then
            AvgCost# = TCost#
          Else
            AvgCost# = AvgCost# / (TOrgQty! + AdjQty!)
          End If
          NewQty# = OrgQty! + AdjQty!
        Else
          AvgCost# = IIf(IsNull(rsInventory("INV ITEM Average Cost")), 0, rsInventory("INV ITEM Average Cost"))
          AvgCost# = (AvgCost# * TOrgQty!) - (TCost# * AdjQty!)
          If (OrgQty! - AdjQty!) <= 0 Then
            AvgCost# = TCost#  'AvgCost#
          Else
            AvgCost# = AvgCost# / (TOrgQty! - AdjQty!)
          End If
          NewQty# = OrgQty! - AdjQty!
        End If
  
          rsInventory("INV ITEM Last Cost") = TCost#
          rsInventory("INV ITEM Average Cost") = AvgCost#
          rsInventory("INV ITEM Qty On Hand") = NewQty#
        rsInventory.Update
        
        rsInventory.Close
        Set rsInventory = Nothing
        
        'Post if necessary
        If rsAdjustment("INV ADJ Post To GL YN") = True Then
          If (AdjustmentType$ = "Increase") Then
  
            '                 Debit   Credit    Source
            '                 -----   ------    ------
            ' Item              X               Iventory GL Acct
            ' Item                       X      Cost of Sale GL Acct
  
  
            If AdjQty! * TCost# = 0 Then
            Else
                SQLstatement = "INSERT INTO [GL Work Detail]"
                SQLstatement = SQLstatement & " ([GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount],[GW TRANSD Number])"
                SQLstatement = SQLstatement & " VALUES ('" & rsAdjustmentDetail("INV ADJD Posting Account") & "'," & AdjQty! * TCost# & ",0," & NewNumber& & ")"
                db.Execute SQLstatement
              
              'rsGLWorkDetail.AddNew
              '  rsGLWorkDetail("GW TRANSD Account") = rsAdjustmentDetail("INV ADJD Posting Account")
              '  rsGLWorkDetail("GW TRANSD Debit Amount") = AdjQty! * TCost#
              '  rsGLWorkDetail("GW TRANSD Credit Amount") = 0
              '  rsGLWorkDetail("GW TRANSD Number") = NewNumber&
              'rsGLWorkDetail.Update
  
                SQLstatement = "INSERT INTO [GL Work Detail]"
                SQLstatement = SQLstatement & " ([GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount],[GW TRANSD Number])"
                SQLstatement = SQLstatement & " VALUES ('" & CostAcct$ & "',0," & AdjQty! * TCost# & "," & NewNumber& & ")"
                db.Execute SQLstatement
              
              'rsGLWorkDetail.AddNew
              '  rsGLWorkDetail("GW TRANSD Account") = CostAcct$
              '  rsGLWorkDetail("GW TRANSD Debit Amount") = 0
              '  rsGLWorkDetail("GW TRANSD Credit Amount") = AdjQty! * TCost#
              '  rsGLWorkDetail("GW TRANSD Number") = NewNumber&
              'rsGLWorkDetail.Update
            End If
  
  
          Else  'Dua Decrease
            '                 Debit   Credit    Source
            '                 -----   ------    ------
            ' Item              X               Cost of Sale Acct
            ' Item                       X      Inventory GL Acct
  
  
            If AdjQty! * TCost# = 0 Then
            Else
                SQLstatement = "INSERT INTO [GL Work Detail]"
                SQLstatement = SQLstatement & " ([GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount],[GW TRANSD Number])"
                SQLstatement = SQLstatement & " VALUES ('" & CostAcct$ & "'," & AdjQty! * TCost# & ",0," & NewNumber& & ")"
                db.Execute SQLstatement
              
              'rsGLWorkDetail.AddNew
              '  rsGLWorkDetail("GW TRANSD Account") = CostAcct$
              '  rsGLWorkDetail("GW TRANSD Debit Amount") = AdjQty! * TCost#
              '  rsGLWorkDetail("GW TRANSD Credit Amount") = 0
              '  rsGLWorkDetail("GW TRANSD Number") = NewNumber&
              'rsGLWorkDetail.Update
  
                SQLstatement = "INSERT INTO [GL Work Detail]"
                SQLstatement = SQLstatement & " ([GW TRANSD Account],[GW TRANSD Debit Amount],[GW TRANSD Credit Amount],[GW TRANSD Number])"
                SQLstatement = SQLstatement & " VALUES ('" & rsAdjustmentDetail("INV ADJD Posting Account") & "',0," & AdjQty! * TCost# & "," & NewNumber& & ")"
                db.Execute SQLstatement
              
              'rsGLWorkDetail.AddNew
              '  rsGLWorkDetail("GW TRANSD Account") = rsAdjustmentDetail("INV ADJD Posting Account")
              '  rsGLWorkDetail("GW TRANSD Debit Amount") = 0
              '  rsGLWorkDetail("GW TRANSD Credit Amount") = AdjQty! * TCost#
              '  rsGLWorkDetail("GW TRANSD Number") = NewNumber&
              'rsGLWorkDetail.Update
            End If
          End If
        Else
          'Don't post to gl
        End If
      End If
      mTotalCost@ = mTotalCost@ + (AdjQty! * TCost#)
    End If
    rsAdjustmentDetail.MoveNext
  Loop
  
  ' post GL entry for this Credit Memo
  If rsAdjustment("INV ADJ Post To GL YN") = True Then
    Dim Success%
    Success% = PostGLWorkDetail(TranDate, NewNumber&, db)
    If Success% = False Then
      MsgBox "An error occurred writing GL Transaction!", , "Error"
      PostAdjustment = False
      Exit Function
    End If
  End If

  PostAdjustment = True
    rsCompany.Close
    Set rsCompany = Nothing
    'rsInventory.Close
    'Set rsInventory = Nothing
    rsAdjustment.Close
    Set rsAdjustment = Nothing
    rsAdjustmentDetail.Close
    Set rsAdjustmentDetail = Nothing
    'rsGLTrans.Close
    'Set rGLTrans = Nothing
    'rsGLWorkDetail.Close
    'Set rsGLWorkDetail = Nothing
   If Currentdb = True Then
    db.Close
    Set db = Nothing
   End If
    
PostAdjustment_Exit:
  Exit Function

UnableToPostAdjustmentHere:
  PostAdjustment = False
      rsCompany.Close
    Set rsCompany = Nothing
    'rsInventory.Close
    'Set rsInventory = Nothing
    rsAdjustment.Close
    Set rsAdjustment = Nothing
    rsAdjustmentDetail.Close
    Set rsAdjustmentDetail = Nothing
    'rsGLTrans.Close
    'Set rGLTrans = Nothing
    'rsGLWorkDetail.Close
    'Set rsGLWorkDetail = Nothing
   If Currentdb = True Then
    db.Close
    Set db = Nothing
   End If
  Exit Function

PostAdjustment_Error:
  Call ErrorLog("Inventory Module", "PostAdjustment", Now, Err.Number, Err.Description, True, db)
  PostAdjustment = False
      rsCompany.Close
    Set rsCompany = Nothing
    'rsInventory.Close
    'Set rsInventory = Nothing
    rsAdjustment.Close
    Set rsAdjustment = Nothing
    rsAdjustmentDetail.Close
    Set rsAdjustmentDetail = Nothing
    'rsGLTrans.Close
    'Set rGLTrans = Nothing
    'rsGLWorkDetail.Close
    'Set rsGLWorkDetail = Nothing
   If Currentdb = True Then
    db.Close
    Set db = Nothing
   End If
  Exit Function

End Function

Function PostProduction(DocumentKey&, Optional db As ADODB.Connection) As Integer
  'On Error GoTo PostProduction_Error
  Dim Currentdb As Boolean
  
  'Dim db As ADODB.Connection
  Currentdb = False
  If db Is Nothing Then
    Set db = New ADODB.Connection
    db.CursorLocation = adUseServer
    db.Open gblADOProvider
    Currentdb = True
  End If
  
  Dim msg$
  Dim title$
  
  Dim rsCompany As ADODB.Recordset
  Set rsCompany = New ADODB.Recordset
  rsCompany.Open "SELECT  [SYS COM GL Post By Date],[SYS COM Sales COGS Acct] " & _
  "FROM [SYS Company]", db, adOpenKeyset, adLockOptimistic, adCmdText
  rsCompany.MoveFirst
'  rsCompany.Open "[SYS Company]", db, adOpenStatic, adLockOptimistic, adCmdTable
'  rsCompany.MoveFirst

  Dim rsProduction As ADODB.Recordset
  Set rsProduction = New ADODB.Recordset
  rsProduction.Open "SELECT [INV PRO Document No],[INV PRO Date],[INV PRO Ext Document No] " & _
  "FROM[INV Production] WHERE [INV PRO Document No]=" & DocumentKey&, db, adOpenKeyset, adLockOptimistic, adCmdText

  ' first lets get the record
  'rsProduction.Index = "PrimaryKey"
  'rsProduction.Seek DocumentKey&
  'rsProduction.Find "[INV PRO Document No]=" & DocumentKey&

  'Post by 1 - system date or 2 - Transaction date?
  Dim PostDate%
  PostDate% = rsCompany("SYS COM GL Post By Date")

  'Set Post Date
  Dim TranDate As Variant
  If PostDate% = 1 Then
    TranDate = DateValue(Format(Now, "Short Date"))
  Else
    TranDate = DateValue(rsProduction("INV PRO Date"))
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
    GoTo UnableToPostProductionHere
  End If

  'On Error GoTo PostProduction_Error

  ' clear any GL Work records
  'Dim cmdtemp As ADODB.Recordset
  'Set cmdtemp = New ADODB.Recordset
  'cmdtemp.Open "DELETE DISTINCTROW * FROM [GL Work Detail]", db, , , adCmdText
  'cmdtemp.Close
  'Set cmdtemp = Nothing
   
  db.Execute "DELETE DISTINCTROW * FROM [GL Work Detail]", , adCmdText
 
  Dim rsProdDetail As ADODB.Recordset
  Set rsProdDetail = New ADODB.Recordset
  rsProdDetail.Open "SELECT [INV PROD Ext Cost],[INV PROD Item ID],[INV PROD Ext Cost]," & _
  "[INV PROD Quantity],[INV PROD Cost] FROM [INV Production Detail] where " & _
  "[INV PROD Document No] = " & DocumentKey&, db, adOpenKeyset, adLockOptimistic, adCmdText
  rsProdDetail.MoveFirst

  Dim rsInventory As ADODB.Recordset
  'Set rsInventory = New ADODB.Recordset
  'rsInventory.Open "[INV Items]", db, adOpenStatic, adLockOptimistic, adCmdTable

  'Dim rsHistory As ADODB.Recordset
  'Set rsHistory = New ADODB.Recordset
  'rsHistory.Open "[INV History]", db, adOpenStatic, adLockOptimistic, adCmdTable

  Dim rsKitItems As ADODB.Recordset

  'Process the Production Changes for the Master Assembly
  
  Dim LastCost#
  Dim QtyonHand#
  Dim ItemAvgCost#
  Dim AvgCost#
  Dim KitID$
  Dim Qty#
  Dim SubID$

  Do While Not rsProdDetail.EOF
    LastCost# = rsProdDetail("INV PROD Ext Cost")
    'Get Qty On Hand
    'rsInventory.Index = "PrimaryKey"
    'rsInventory.Seek rsProdDetail("INV PROD Item ID")
      Set rsInventory = New ADODB.Recordset
      rsInventory.Open "SELECT [INV ITEM Id],[INV ITEM Qty On Hand],[INV ITEM Average Cost] " & _
      "FROM [INV Items] WHERE [INV ITEM Id]='" & rsProdDetail("INV PROD Item ID") & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
      rsInventory.MoveFirst
    'rsInventory.MoveFirst
    'rsInventory.Find "[INV ITEM Id]='" & rsProdDetail("INV PROD Item ID") & "'"
    
    If rsInventory.RecordCount = 0 Then
      'This is a problem
    Else
      QtyonHand# = IIf(IsNull(rsInventory("INV ITEM Qty On Hand")), 0, rsInventory("INV ITEM Qty On Hand"))
      ItemAvgCost# = IIf(IsNull(rsInventory("INV ITEM Average Cost")), 0, rsInventory("INV ITEM Average Cost"))
    End If
    If QtyonHand# < 0 Then QtyonHand# = 0
    AvgCost# = ((ItemAvgCost# * QtyonHand#) + (rsProdDetail("INV PROD Ext Cost")))
    If QtyonHand# + rsProdDetail("INV PROD Quantity") <= 0 Then
      AvgCost# = rsProdDetail("INV PROD Cost")
    Else
      AvgCost# = AvgCost# / (QtyonHand# + rsProdDetail("INV PROD Quantity"))
    End If
      rsInventory("INV ITEM Average Cost") = DropAllBut2(CCur(AvgCost#))
      rsInventory("INV ITEM Qty On Hand") = rsInventory("INV ITEM Qty On Hand") + rsProdDetail("INV PROD Quantity")
    rsInventory.Update

    'On Error Resume Next
    Set rsKitItems = New ADODB.Recordset
    rsKitItems.Open "SELECT [INV KIT Qty],[INV KIT Sub Item ID] FROM [INV Kit Items 2] where [INV KIT Item ID] = '" & rsProdDetail("INV PROD Item ID") & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    'rsKitItems.MoveLast
    'rsKitItems.MoveFirst
    If rsKitItems.RecordCount = 0 Then
    Else
      Do While Not rsKitItems.EOF
        Qty# = IIf(IsNull(rsKitItems("INV KIT Qty")), 0, rsKitItems("INV KIT Qty"))
        SubID$ = IIf(IsNull(rsKitItems("INV KIT Sub Item ID")), "", rsKitItems("INV KIT Sub Item ID"))
        'rsInventory.Index = "PrimaryKey"
        'rsInventory.Seek SubID$
        
        rsInventory.Close
        Set rsInventory = Nothing
        
        Set rsInventory = New ADODB.Recordset
        rsInventory.Open "SELECT [INV ITEM Id],[INV ITEM Qty On Hand],[INV ITEM Average Cost] " & _
        "FROM [INV Items] WHERE [INV ITEM Id]='" & SubID$ & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
        rsInventory.MoveFirst
        'rsInventory.Find "[INV ITEM Id]='" & SubID$ & "'"
        If rsInventory.RecordCount = 0 Then
         'There is a problem
        Else
          QtyonHand# = IIf(IsNull(rsInventory("INV ITEM Qty On Hand")), 0, rsInventory("INV ITEM Qty On Hand"))
          AvgCost# = IIf(IsNull(rsInventory("INV ITEM Average Cost")), 0, rsInventory("INV ITEM Average Cost"))
            rsInventory("INV ITEM Qty On Hand") = QtyonHand# - (Qty# * rsProdDetail("INV PROD Quantity"))
          rsInventory.Update
        End If
        
        'Write an inventory history record
        SQLstatement = "INSERT INTO [INV History]"
        SQLstatement = SQLstatement & " ([INV HIST Document No],[INV HIST Type],[INV HIST Date],[INV HIST Item ID],[INV HIST Quantity],[INV HIST Cost])"
        SQLstatement = SQLstatement & " VALUES ('" & rsProduction("INV PRO Ext Document No") & "','Production',#" & rsProduction("INV PRO Date") & "#,'" & SubID$ & "'," & (Qty# * rsProdDetail("INV PROD Quantity")) * -1 & "," & AvgCost# & ")"
        db.Execute SQLstatement
        
        'rsHistory.AddNew
        '  rsHistory("INV HIST Document No") = rsProduction("INV PRO Ext Document No")
        '  rsHistory("INV HIST Type") = "Production"
        '  rsHistory("INV HIST Date") = rsProduction("INV PRO Date")
        '  rsHistory("INV HIST Item ID") = SubID$
        '  rsHistory("INV HIST Quantity") = (Qty# * rsProdDetail("INV PROD Quantity")) * -1
        '  rsHistory("INV HIST Cost") = AvgCost#
        'rsHistory.Update


        rsKitItems.MoveNext
      Loop
    End If
    rsProdDetail.MoveNext
    
    rsKitItems.Close
    Set rsKitItems = Nothing
    If rsInventory Is Nothing Then
    Else
        rsInventory.Close
        Set rsInventory = Nothing
    End If
  Loop

  PostProduction% = True
  rsCompany.Close
  Set rsCompany = Nothing
  'rsHistory.Close
  'Set rsHistory = Nothing
  'rsInventory.Close
  'Set rsInventory = Nothing
  'rsKitItems.Close
  'Set rsKitItems = Nothing
  rsProdDetail.Close
  Set rsProdDetail = Nothing
'  rs.production.Close
  Set rsProduction = Nothing
   If Currentdb = True Then
    db.Close
    Set db = Nothing
   End If
  Exit Function

UnableToPostProductionHere:
  PostProduction% = False
    rsCompany.Close
  Set rsCompany = Nothing
  'rsHistory.Close
  'Set rsHistory = Nothing
  'rsInventory.Close
  'Set rsInventory = Nothing
  'rsKitItems.Close
  'Set rsKitItems = Nothing
  rsProdDetail.Close
  Set rsProdDetail = Nothing
  rs.production.Close
  Set rsProduction = Nothing
   If Currentdb = True Then
    db.Close
    Set db = Nothing
   End If
  Exit Function

PostProduction_Error:
  Call ErrorLog("Inventory Module", "PostProduction", Now, Err.Number, Err.Description, True, db)
  PostProduction% = False
    rsCompany.Close
  Set rsCompany = Nothing
  'rsHistory.Close
  'Set rsHistory = Nothing
  'rsInventory.Close
  'Set rsInventory = Nothing
  'rsKitItems.Close
  'Set rsKitItems = Nothing
  rsProdDetail.Close
  Set rsProdDetail = Nothing
  rs.production.Close
  Set rsProduction = Nothing
   If Currentdb = True Then
    db.Close
    Set db = Nothing
   End If
  Exit Function

End Function

