Attribute VB_Name = "Validation_Module"

Function ValidAccount%(Acct$)

  'On Error GoTo ValidAccount_Error

  Dim db As ADODB.Connection
  Set db = New Connection
  db.CursorLocation = adUseClient
  db.Open gblADOProvider
  Dim rs As ADODB.Recordset
  rs.Open "SELECT * FROM [GL Chart Of Accounts] WHERE [GL COA Account No] = '" & Acct$ & "'", db, adOpenStatic, adLockOptimistic, adCmdText
  
  'On Error Resume Next

  rs.MoveFirst
  If (rs.BOF And rs.EOF) Then
    ValidAccount% = False
  Else
    If rs("GL COA Asset Type") = "Report Title" Then
      ValidAccount% = False
    Else
      ValidAccount% = True
    End If
  End If

 rs.Close
 Set rs = Nothing
 db.Close
 Set db = Nothing
  
  Exit Function
ValidAccount_Error:
  Call LogError("Validation Routines", "ValidAccount", Now, Err, Error, True)
  Resume Next

End Function

Function ValidCustomer(CustomerID$) As Integer

  'On Error GoTo ValidCustomer_Error

  Dim db As ADODB.Connection
  Set db = New Connection
  db.CursorLocation = adUseClient
  db.Open gblADOProvider
  Dim rs As ADODB.Recordset
  rs.Open "SELECT * FROM [AR Customer] WHERE [AR CUST Customer ID] = '" & CustomerID$ & "'", db, adOpenStatic, adLockOptimistic, adCmdText
  
  'On Error Resume Next

  rs.MoveFirst
  If (rs.BOF And rs.EOF) Then
    ValidCustomer% = False
  Else
    ValidCustomer% = True
  End If
  
  rs.Close
  Set rs = Nothing
  db.Close
  Set db = Nothing
  Exit Function
ValidCustomer_Error:
  Call LogError("Validation Routines", "ValidCustomer", Now, Err, Error, True)
  Resume Next

End Function
Function ValidEmployee(EmpID$) As Integer

  'On Error GoTo ValidEmployee_Error

  Dim db As ADODB.Connection
  Set db = New Connection
  db.CursorLocation = adUseClient
  db.Open gblADOProvider
  Dim rs As ADODB.Recordset
  rs.Open "SELECT * FROM [EMP Employees] WHERE [EMP ID] = '" & EmpID$ & "'", db, adOpenStatic, adLockOptimistic, adCmdText

  'On Error Resume Next
  rs.MoveFirst
  If (rs.BOF And rs.EOF) Then
    ValidEmployee% = False
  Else
    ValidEmployee% = True
  End If

  rs.Close
  Set rs = Nothing
  db.Close
  Set db = Nothing
  
  Exit Function
ValidEmployee_Error:
  Call LogError("Validation Routines", "ValidEmployee", Now, Err, Error, True)
  Resume Next

End Function


Function ValidItem(ItemID$) As Integer

  'On Error GoTo ValidItem_Error
    
  Dim db As ADODB.Connection
  Set db = New Connection
  db.CursorLocation = adUseClient
  db.Open gblADOProvider
  Dim rs As ADODB.Recordset
  rs.Open "SELECT * FROM [INV Items] WHERE [INV Item ID] = '" & ItemID$ & "'", db, adOpenStatic, adLockOptimistic, adCmdText

  'On Error Resume Next
  rs.MoveFirst
  If (rs.BOF And rs.EOF) Then
    ValidItem% = False
  Else
    ValidItem% = True
  End If

  rs.Close
  Set rs = Nothing
  db.Close
  Set db = Nothing
  Exit Function
ValidItem_Error:
  Call LogError("Validation Routines", "ValidItem", Now, Err, Error, True)
  rs.Close
  Set rs = Nothing
  db.Close
  Set db = Nothing
  Exit Function

End Function


Function ValidProject(ProjID$) As Integer

  'On Error GoTo ValidProject_Error

  Dim db As ADODB.Connection
  Set db = New Connection
  db.CursorLocation = adUseClient
  db.Open gblADOProvider
  Dim rs As ADODB.Recordset
  rs.Open "SELECT * FROM [PROJ Projects] WHERE [PROJ ID] = '" & ProjID$ & "'", db, adOpenStatic, adLockOptimistic, adCmdText

  'On Error Resume Next
  rs.MoveFirst
  If (rs.BOF And rs.EOF) Then
    ValidProject% = False
  Else
    ValidProject% = True
  End If

  rs.Close
  Set rs = Nothing
  db.Close
  Set db = Nothing
  Exit Function
ValidProject_Error:
  Call LogError("Validation Routines", "ValidProject", Now, Err, Error, True)
  Resume Next

End Function

Function ValidTax(TaxID$) As Integer

  'On Error GoTo ValidTax_Error

  Dim db As ADODB.Connection
  Set db = New Connection
  db.CursorLocation = adUseClient
  db.Open gblADOProvider
  Dim rs As ADODB.Recordset
  rs.Open "SELECT * FROM [SYS Tax] WHERE [SYS TAX ID] = '" & TaxID$ & "'", db, adOpenStatic, adLockOptimistic, adCmdText

  'On Error Resume Next
  rs.MoveFirst
  If (rs.BOF And rs.EOF) Then
    ValidTax% = False
  Else
    ValidTax% = True
  End If
  
  rs.Close
  Set rs = Nothing
  db.Close
  Set db = Nothing
  
  Exit Function
ValidTax_Error:
  Call LogError("Validation Routines", "ValidTax", Now, Err, Error, True)
  Resume Next
  
End Function

Function ValidTaxGroup(TaxGrpID$) As Integer

  'On Error GoTo ValidTaxGroup_Error

  Dim db As ADODB.Connection
  Set db = New Connection
  db.CursorLocation = adUseClient
  db.Open gblADOProvider
  Dim rs As ADODB.Recordset
  rs.Open "SELECT * FROM [SYS Tax Group] WHERE [SYS TAXGRP ID] = '" & TaxGrpID$ & "'", db, adOpenStatic, adLockOptimistic, adCmdText

  'On Error Resume Next
  rs.MoveFirst
  If (rs.BOF And rs.EOF) Then
    ValidTaxGroup% = False
  Else
    ValidTaxGroup% = True
  End If

  rs.Close
  Set rs = Nothing
  db.Close
  Set db = Nothing
  
  Exit Function
ValidTaxGroup_Error:
  Call LogError("Validation Routines", "ValidTaxGroup", Now, Err, Error, True)
  Resume Next

End Function

Function ValidVendor(VendorID$) As Integer

  'On Error GoTo ValidVendor_Error

  Dim db As ADODB.Connection
  Set db = New Connection
  db.CursorLocation = adUseClient
  db.Open gblADOProvider
  Dim rs As ADODB.Recordset
  rs.Open "SELECT * FROM [AP Vendor] WHERE [AP VEN ID] = '" & VendorID$ & "'", db, adOpenStatic, adLockOptimistic, adCmdText

  'On Error Resume Next
  rs.MoveFirst
  If (rs.BOF And rs.EOF) Then
    ValidVendor% = False
  Else
    ValidVendor% = True
  End If
  
  rs.Close
  Set rs = Nothing
  db.Close
  Set db = Nothing
  
  Exit Function
ValidVendor_Error:
  Call LogError("Validation Routines", "ValidVendor", Now, Err, Error, True)
  Resume Next

End Function

Sub VerifyPeriod(TranDate As Variant, PeriodToPost%, PeriodClosed%, Optional db As ADODB.Connection)
Dim Currentdb As Boolean

  ReDim Period(1 To 14) As Variant
  Dim HighestPeriod%
  Dim X%
  
  'On Error Resume Next

  'Dim db As ADODB.Connection
  Currentdb = False
  If db Is Nothing Then
    Set db = New Connection
    db.CursorLocation = adUseServer
    db.Open gblADOProvider
    Currentdb = True
  End If
  
  Dim SQLstatement As String
    SQLstatement = "[SYS COM P1 Date],[SYS COM P1 Closed]"
  For X% = 2 To 13
    SQLstatement = SQLstatement & ",[SYS COM P" & Trim$(Str(X%)) & " Date],[SYS COM P" & Trim$(Str(X%)) & " Closed]"
  Next X%
  SQLstatement = SQLstatement & ",[SYS COM Fiscal End Date]"
  SQLstatement = "SELECT " & SQLstatement & " FROM [SYS Company]"
  
  Dim rsCompany As ADODB.Recordset
  Set rsCompany = New Recordset
  rsCompany.Open SQLstatement, db, adOpenForwardOnly, adLockOptimistic, adCmdText '<<<---3 seconds

  rsCompany.MoveFirst
  
  For X% = 1 To 13
    Period(X%) = IIf(IsNull(rsCompany("SYS COM P" & Trim$(Str(X%)) & " Date")), Null, rsCompany("SYS COM P" & Trim$(Str(X%)) & " Date"))
  Next X%
  Period(14) = IIf(IsNull(rsCompany("SYS COM Fiscal End Date")), DateValue(Format(DateSerial(Val(Format(TranDate, "yyyy")) + 1, 1, 1), "Short Date")), DateAdd("d", 1, rsCompany("SYS COM Fiscal End Date")))
    
  HighestPeriod% = -1
  For X% = 2 To 13
    If IsNull(Period(X%)) Then
      HighestPeriod% = X% - 1
      Exit For
    End If
  Next X%
  If HighestPeriod% = -1 Then HighestPeriod% = 13
EndCalculateHighestPeriod:

    ' Test Period to Post
  If DateValue(TranDate) >= DateValue(Period(14)) Then
    PeriodToPost% = 14
    GoTo EndOfTestPeriodToPost
  End If

  If HighestPeriod% = 1 Then ' Or DateValue(TranDate) < DateValue(Period(2)) Then
    PeriodToPost% = 1
    GoTo EndOfTestPeriodToPost
  End If

  If DateValue(TranDate) < DateValue(Period(1)) Then
    PeriodToPost% = 0
    GoTo EndOfTestPeriodToPost
  End If

  PeriodToPost% = -1
  For X% = 1 To HighestPeriod% - 1
    If DateValue(TranDate) >= DateValue(Period(X%)) And DateValue(TranDate) < DateValue(Period(X% + 1)) Then
      PeriodToPost% = X%
      Exit For
    End If
  Next X%
EndOfTestPeriodToPost:
  If PeriodToPost% = -1 Then PeriodToPost% = HighestPeriod%
    ' End of Test Period to Post

  'Determine if period to post is closed
  If PeriodToPost% = 14 Then
      MsgBox "The input date [" & TranDate & "] value is exceed the FISCAL END DATE," & vbCr & " it will be posted to the end of period [" & Period(14) & "]", vbCritical, "Warning"
      PeriodToPost% = PeriodToPost% - 1
  End If
  If PeriodToPost% = 0 Then Exit Sub   '<<<---2 seconds
  PeriodClosed% = IIf(IsNull(rsCompany("SYS COM P" & Trim$(Str(PeriodToPost%)) & " Closed")), 0, rsCompany("SYS COM P" & Trim$(Str(PeriodToPost%)) & " Closed"))
  PeriodClosed% = PeriodClosed% '* -1

rsCompany.Close
Set rsCompany = Nothing

   If Currentdb = True Then
    db.Close
    Set db = Nothing
   End If
End Sub

