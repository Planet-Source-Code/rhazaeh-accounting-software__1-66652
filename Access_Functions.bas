Attribute VB_Name = "Primary_Function"
Option Explicit

Public Function CheckNewDB(ADOprimaryrs As ADODB.Recordset, FormType As String) As Boolean
  With ADOprimaryrs
    If .RecordCount = 0 Then
        MsgBox FormType & " is empty. Creating new " & FormType & ".", vbInformation, "Information"
        
        CheckNewDB = True
        Exit Function
    End If
  End With
  CheckNewDB = False
End Function

Public Function GetcheckNumber(db As ADODB.Connection, BankAcctID As String) As String
  Dim rsBank As ADODB.Recordset
  Set rsBank = New ADODB.Recordset
  rsBank.Open "SELECT * FROM [Bank Accounts] WHERE [BANK ACCT ID]='" & BankAcctID & "'", db, adOpenStatic, adLockOptimistic
  If rsBank.RecordCount = 0 Then
     MsgBox "There is an error on Bank setup!", vbCritical, "Critical Error"
     GetcheckNumber = "Error"
     Exit Function
  End If
  GetcheckNumber = rsBank("BANK ACCT Next Check No")
  rsBank.Close
  Set rsBank = Nothing
End Function

Public Function DataDelete(ADOprimaryrs As ADODB.Recordset, frm As Form, UseOrderBy As Boolean) As Boolean
ShowStatus True
On Error GoTo DeleteErr
If ADOprimaryrs.RecordCount > 0 Then
  With ADOprimaryrs
        If .AbsolutePosition = 1 And UseOrderBy = False Then
           MsgBox "Deleting first record is not allowed. Deleting instruction is cancelled", vbInformation, "Information"
           DataDelete = False
           ShowStatus False
           Exit Function
        End If

        If .EditMode <> 0 Then
          .CancelUpdate
        End If
        'MsgBox .EditMode & "   " & "adEditNone = " & adEditNone & " adEditInProgress = " & adEditInProgress & " adEditAdd = " & adEditAdd & " adEditDelete = " & adEditDelete
        .Delete
        On Error GoTo exit_sub
        If .EOF Then
            .MoveFirst
        Else
            .MovePrevious
        End If
  End With
  DataDelete = True
End If

  ShowStatus False
  Exit Function
DeleteErr:
  ShowStatus False
  DataDelete = False
exit_sub:
  DataDelete = True
End Function


Public Function DataGridKnownError(DataError As Integer) As Boolean
    Select Case DataError
    Case 6153, 7007
    'this is a known error but does not affect the result so i disable the messageBox
    '     MsgBox "DataError : " & DataError & " This Error gives no Harm... but if someone knows how " & vbCr & "to fix this please e-mail to erricoe@TBS.com" & vbCr & " Under Subject: e-mail Razi...He needs my Help", vbCritical, "Known Bugs"
         DataGridKnownError = True
    Case 7011
    'occured when deleting process is cancell
         DataGridKnownError = True
    Case 6152
    '---
         DataGridKnownError = False
    Case Else
         DataGridKnownError = False
    End Select

End Function


Public Function UnloadForm(ADOprimaryrs As ADODB.Recordset) As Integer
On Error Resume Next
If CloseAllActive = True Or ADOprimaryrs.RecordCount = 0 Then
    ADOprimaryrs.CancelUpdate
    ADOprimaryrs.Close
    Set ADOprimaryrs = Nothing
    UnloadForm = 0
    Exit Function
Else
    With ADOprimaryrs
    If .EditMode <> adEditNone And .EditMode <> adEditAdd Then
    Dim CreateOrder As Integer
    CreateOrder = MsgBox("Attempting to close the application. " & vbCr & "Would you like to update the data?", vbYesNoCancel, "Exiting")
        If CreateOrder = vbYes Then
           'If .Status = adRecModified Or .Status = adRecNew Then .Update
           .Update
           .Close
           Set ADOprimaryrs = Nothing
           UnloadForm = 0
        ElseIf CreateOrder = vbCancel Then
           UnloadForm = 1
        Else
           'If .Status = adRecModified Or .Status = adRecNew Then .CancelUpdate
           .CancelUpdate
           ADOprimaryrs.Close
           Set ADOprimaryrs = Nothing
           UnloadForm = 0
        End If
    Else
           .CancelUpdate
           ADOprimaryrs.Close
           Set ADOprimaryrs = Nothing
           UnloadForm = 0
    End If
    
    End With
End If
End Function

Public Function CheckDocument(SQLstatement As String, db As ADODB.Connection, Optional ShowMessage As Boolean, Optional txtCallType As TextBox, Optional lblCallType As String) As Boolean
ShowStatus True
Dim dbCnn As ADODB.Connection
Dim cbRS As ADODB.Recordset
'Dim i As Integer

  'Set dbCnn = New ADODB.Connection
  'dbCnn.CursorLocation = adUseClient
  'dbCnn.Open gblADOProvider
  'Debug.Print SQLstatement
  
  Set cbRS = New ADODB.Recordset
  With cbRS
    'MsgBox SQLStatement
    .Open SQLstatement, db, adOpenKeyset, adLockReadOnly, adCmdText
    If .RecordCount > 0 Then
       'trap this error
       If ShowMessage = True Then
            MsgBox "The value is already in used!", vbCritical, "Error"
       End If
       CheckDocument = False
       GoTo FalseFlag
    End If
  .Close
  End With
CheckDocument = True

Dim Response As Integer

If txtCallType Is Nothing Then
Else
    If lblCallType = "" Then
        lblCallType = fMainForm.ActiveForm.lblfields(txtCallType.Index).Caption
    End If
    Response = MsgBox("This is a new input for " & lblCallType & " , would you like to add it to the database?", vbYesNo, "Information")
    If Response = vbYes Then
    
        Select Case lblCallType '
        Case "Bank Account"
            frm_SYS_Setup_Chart_Of_Accounts.CallByUserCOA txtCallType.Text, fMainForm.ActiveForm.lblfields(txtCallType.Index).Caption
            frm_SYS_Setup_Chart_Of_Accounts.ZOrder 0
            MsgBox "New data have been transferred to CHART OF ACCOUNT, but you have to add more data and save it.", vbInformation, "Information"
        Case "COA"
            frm_SYS_Setup_Chart_Of_Accounts.CallByUserCOA txtCallType.Text, fMainForm.ActiveForm.lblfields(txtCallType.Index).Caption
            frm_SYS_Setup_Chart_Of_Accounts.ZOrder 0
            MsgBox "New data have been transferred to CHART OF ACCOUNT, but you have to add more data and save it.", vbInformation, "Information"
        Case "Customer ID"
            frm_AR_Customer.CallByUserCust txtCallType.Text
            frm_AR_Customer.ZOrder 0
            MsgBox "New data have been transferred to CUSTOMER SETUP, but you have to add more data and save it.", vbInformation, "Information"
        Case "Shipping ID"
            frm_AR_Cust_Ship_To.CallByUserShip txtCallType.Text
            frm_AR_Cust_Ship_To.ZOrder 0
            MsgBox "New data have been transferred to SHIP TO SETUP, but you have to add more data and save it.", vbInformation, "Information"
        End Select
        txtCallType.Text = " "
    Else
        txtCallType.Text = " "
    End If
End If
FalseFlag:
Set cbRS = Nothing
'dbCnn.Close
'Set dbCnn = Nothing
ShowStatus False
End Function

'for each control whose input you wish to validate, just put something like this
'in the KeyPress event of the control-->>keyResponse=CtrlValidate(Keyascii, "$.0123456789")
'Doing so will filter out any undesired keys that go to the control, accepting
'only the keys defined by the second parameter. In this case, that parameter
'("$.0123456789") defines characters that are valid for a currency but put this after the code
'    If keyResponse = False Then
'       KeyAscii = 0
'    End If

Public Function CtrlValidate(KeyIn As Integer, ValidateString As String) As Boolean
Dim ValidateList As String
Dim KeyOut As Integer

If KeyIn = 8 Or KeyIn = 9 Then
    CtrlValidate = True
    Exit Function
End If
If InStr(1, ValidateString, Chr(KeyIn), 1) > 0 Then
   CtrlValidate = True
Else
   CtrlValidate = False
   Beep
End If
End Function

Public Function CheckCombo(SourceCombo As ComboBox, Optional FieldName As String, Optional TableName As String, Optional db As ADODB.Connection, Optional MsgLoad As Boolean) As Boolean
On Error GoTo exit_function
Dim i As Integer
Dim Response As Integer

CheckCombo = True
  
  For i = 0 To SourceCombo.ListCount - 1
    If SourceCombo.Text = SourceCombo.List(i) Then
        CheckCombo = False
    End If
  Next
  
If CheckCombo = True And MsgLoad = True And SourceCombo.Text <> "" Then
    Dim ADOnewRS As ADODB.Recordset
    Response = MsgBox("This is a new input for " & fMainForm.ActiveForm.lblfields(SourceCombo.Index) & " , would you like to add it to the database?", vbYesNo, "Information")
    If Response = vbYes Then
        Select Case TableName
        Case "[SYS Tax Group]"
            frm_SYS_Setup_Tax_Group.CallByUser SourceCombo.Text
            MsgBox "New data have been transferred to TAX GROUP, but you have to add more data and save it." & vbCr & "Then the data will be available to use after you click the refresh button", vbInformation, "Information"
            SourceCombo.Text = SourceCombo.List(0)
        Case "[LIST Payment Terms]"
            frm_LIST_Payment_Terms.CallByUserPay SourceCombo.Text
            MsgBox "New data have been transferred to PAY TERMS, but you have to add more data and save it." & vbCr & "Then the data will be available to use after you click the refresh button", vbInformation, "Information"
            SourceCombo.Text = SourceCombo.List(0)
        'Case "[EMP Employees]"
        '    frm_SYS_Setup_Employee.CallByUserEmpID SourceCombo.Text
        '    MsgBox "New data have been transferred to PAY TERMS, but you have to add more data and save it." & vbCr & "Then the data will be available to use after you click the refresh button", vbInformation, "Information"
        '    SourceCombo.Text = SourceCombo.List(0)
        Case "[LIST Payment Methods]"
           Set ADOnewRS = New ADODB.Recordset
           With ADOnewRS
                .Open "SELECT [LIST PAY Method],[Payment Terms] FROM [LIST Payment Methods] WHERE " & FieldName & "='" & SourceCombo.Text & "' AND [Payment Terms]='" & fMainForm.ActiveForm.cbPurchase(5).Text & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
                If .RecordCount = 0 Then
                    .AddNew
                        ![LIST PAY Method] = SourceCombo.Text & ""
                        ![Payment Terms] = fMainForm.ActiveForm.cbPurchase(5).Text
                        SourceCombo.List(i) = SourceCombo.Text
                    .Update
                End If
                .Close
           End With
           Set ADOnewRS = Nothing
        Case Else 'Customer Type, Item Categories,Payment Methods,Shipping Methods,
                  'Vendor Type, Recurring Type
           Set ADOnewRS = New ADODB.Recordset
           With ADOnewRS
                .Open "SELECT " & FieldName & " FROM " & TableName & " WHERE " & FieldName & "='" & SourceCombo.Text & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
                If .RecordCount = 0 Then
                    .AddNew
                        .Fields("" & StripBrackets(FieldName) & "") = SourceCombo.Text & ""
                        SourceCombo.List(i) = SourceCombo.Text
                    .Update
                End If
                .Close
           End With
           Set ADOnewRS = Nothing
        End Select
    Else
        SourceCombo.Text = SourceCombo.List(0)
    End If
ElseIf CheckCombo = True And MsgLoad = False Then
    SourceCombo.Text = SourceCombo.List(0)
ElseIf CheckCombo = False Then
    Select Case TableName
    Case "[LIST Shipping Methods]"
        Set ADOnewRS = New ADODB.Recordset 'txtfields(27)
        ADOnewRS.Open " Select [LIST SHIP Charge] FROM [LIST Shipping Methods] WHERE [LIST SHIP Method]='" & SourceCombo.Text & "'", db, adOpenKeyset, adLockReadOnly, adCmdText
          If fMainForm.ActiveForm.txtFields(23).Enabled = True Then fMainForm.ActiveForm.txtFields(23).SetFocus
          If ADOnewRS![LIST SHIP Charge] > 0 Then
            fMainForm.ActiveForm.txtFields(23) = "Yes"
            fMainForm.ActiveForm.txtFields(27) = FormatCurr(ADOnewRS![LIST SHIP Charge])
          Else
            fMainForm.ActiveForm.txtFields(23) = "No"
            fMainForm.ActiveForm.txtFields(27) = "$0.00"
          End If
        ADOnewRS.Close
        Set ADOnewRS = Nothing
    Case "[SYS Tax Group]"
        Set ADOnewRS = New ADODB.Recordset 'txtfields(27)
        ADOnewRS.Open " Select [SYS TAXGRPD Tax ID] FROM [SYS Tax Group Detail] WHERE [SYS TAXGRPD Group ID]='" & SourceCombo.Text & "'", db, adOpenKeyset, adLockReadOnly, adCmdText
        Dim CountPercent As Double
        CountPercent = 0
          If ADOnewRS.RecordCount > 0 Then
                Dim ADOSecondRS As ADODB.Recordset
                
                Do While Not ADOnewRS.EOF
                  Set ADOSecondRS = New ADODB.Recordset
                  ADOSecondRS.Open " Select [SYS TAX Percent] FROM [SYS Tax] WHERE [SYS TAX ID]='" & ADOnewRS![SYS TAXGRPD Tax ID] & "'", db, adOpenKeyset, adLockReadOnly, adCmdText
                    CountPercent = CountPercent + CDbl(ADOSecondRS![SYS TAX Percent])
                  ADOSecondRS.Close
                  Set ADOSecondRS = Nothing
                  ADOnewRS.MoveNext
                Loop
          End If
                  fMainForm.ActiveForm.txtFields(29).SetFocus
                  If CountPercent > 0 Then
                    fMainForm.ActiveForm.txtFields(29) = Format(CountPercent, "00.00")
                  Else
                    fMainForm.ActiveForm.txtFields(29) = "00.00"
                  End If
        ADOnewRS.Close
        Set ADOnewRS = Nothing
    End Select
End If

exit_function:
End Function

Public Function SumRecord(Searchee As String, TableName As String, db As ADODB.Connection, Optional WhereCriteria As String) As Variant
'This function use to SUM all of the selected record. It uses the ADO Access Method.
Dim Currentdb As Boolean
'
Exit Function
  'Dim db As ADODB.Connection
  Currentdb = False
  If db Is Nothing Then
    Set db = New ADODB.Connection
    db.CursorLocation = adUseServer
    db.Open gblADOProvider
    Currentdb = True
  End If
    'Then Open the Recordset
    Dim ADOprimaryrs As ADODB.Recordset
    Set ADOprimaryrs = New ADODB.Recordset
    'Execute the Query
    If IsMissing(WhereCriteria) Then
        'ADOprimaryRS.Source = "Select " & Searchee & " From " & TableName
        'Set ADOprimaryRS.ActiveConnection = con
        ADOprimaryrs.Open "Select " & Searchee & " From " & TableName, db, adOpenKeyset, adLockReadOnly, adCmdText
    Else
        ADOprimaryrs.Open "Select " & Searchee & " From " & TableName & " Where " & WhereCriteria, db, adOpenKeyset, adLockReadOnly, adCmdText
        'ADOprimaryRS.Requery
        'ADOprimaryRS.Source = "Select " & Searchee & " From [" & TableName & "] Where " & WhereCriteria
        'Debug.Print "Select " & Searchee & " From [" & TableName & "] Where " & WhereCriteria
        'Set ADOprimaryRS.ActiveConnection = con
        'ADOprimaryRS.Open
    End If
        
    SumRecord = 0
    If Not ADOprimaryrs.EOF Then ADOprimaryrs.MoveFirst
    If ADOprimaryrs.RecordCount > 0 Then ADOprimaryrs.MoveFirst
    
    While Not ADOprimaryrs.EOF
        If Not IsNull(ADOprimaryrs.Fields(0).Value) Then
            SumRecord = SumRecord + CDbl(ADOprimaryrs.Fields(0).Value)
        End If
        ADOprimaryrs.MoveNext
    Wend

    ADOprimaryrs.Close
    Set ADOprimaryrs = Nothing
    
   If Currentdb = True Then
    db.Close
    Set db = Nothing
   End If
End Function

Public Function CountRecord(Searchee As String, TableName As String, db As ADODB.Connection, Optional WhereCriteria As String) As Variant
   'This function duplicates the CountRecord function in access. It is written using the global connect strings set-up in
   'the Sub-Main Function in the Main Access Module. It uses the ADO Access Method.
   'Open the Connection
Dim Currentdb As Boolean
  Currentdb = False
  If db Is Nothing Then
    Set db = New ADODB.Connection
    db.CursorLocation = adUseServer
    db.Open gblADOProvider
    Currentdb = True
  End If
    'Open the Recordset
    Dim ADOprimaryrs As ADODB.Recordset
    Set ADOprimaryrs = New ADODB.Recordset
    'Set ADOprimaryRS.ActiveConnection = con
    'filter and recordcount also perform the same way
    If IsMissing(WhereCriteria) Then
        ADOprimaryrs.Open "Select Count (" & Searchee & ")  From " & TableName, db, adOpenKeyset, adLockReadOnly, adCmdText
    Else
        ADOprimaryrs.Open "Select Count (" & Searchee & ")  From " & TableName & " Where " & WhereCriteria, db, adOpenKeyset, adLockReadOnly, adCmdText
    End If
    'Debug.Print "Select Count (" & Searchee & ")  From " & TableName & " Where " & WhereCriteria
    'If IsNull(ADOprimaryRS.Fields(0).Value) Then
    '    CountRecord = 0
    'Else
        'CountRecord = ADOprimaryrs.RecordCount 'ADOprimaryRS.Fields(0).Value
    'End If
    If IsNull(ADOprimaryrs.Fields(0).Value) Then
        CountRecord = 0
    Else
        CountRecord = ADOprimaryrs.Fields(0).Value
    End If
    
    ADOprimaryrs.Close
    Set ADOprimaryrs = Nothing
    
   If Currentdb = True Then
    db.Close
    Set db = Nothing
   End If
End Function

Public Function LookRecord(Searchee As String, TableName As String, db As ADODB.Connection, Optional WhereCriteria As String) As Variant
    'This function duplicates the LookRecord function in access. It is written using the global connect strings set-up in
    'the Sub-Main Function in the Main Access Module. It uses the ADO Access Method.
    'Open the Connection
Dim Currentdb As Boolean
  Currentdb = False
  If db Is Nothing Then
    Set db = New ADODB.Connection
    db.CursorLocation = adUseServer
    db.Open gblADOProvider
    Currentdb = True
  End If
    'Open the Recordset
    Dim ADOprimaryrs As ADODB.Recordset
    Set ADOprimaryrs = New ADODB.Recordset
    'Set ADOprimaryRS.ActiveConnection = con
            
    If WhereCriteria = "" Then
         ADOprimaryrs.Open "Select (" & Searchee & ") From " & TableName, db, adOpenKeyset, adLockReadOnly, adCmdText
    Else
         ADOprimaryrs.Open "Select " & Searchee & " From " & TableName & " Where " & WhereCriteria, db, adOpenKeyset, adLockReadOnly, adCmdText
    End If
    
    If ADOprimaryrs.RecordCount = 0 Then
        MsgBox "The Database is Empty.", vbCritical, "Error"
        LookRecord = " "
        GoTo EmptyDB
    End If
    If IsNull(ADOprimaryrs.Fields("" & StripBrackets(Searchee) & "").Value) Then
        LookRecord = " "
    Else
        LookRecord = ADOprimaryrs.Fields("" & StripBrackets(Searchee) & "").Value
    End If
    
EmptyDB:
    ADOprimaryrs.Close
    Set ADOprimaryrs = Nothing
    
   If Currentdb = True Then
    db.Close
    Set db = Nothing
   End If
End Function

Public Function NZ(Check1 As Variant, Optional Check2 As Variant)
    'This function duplicates the NZ function in access.
    If IsNull(Check1) Then
        If IsMissing(Check2) Then NZ = 0 Else NZ = Check2
    Else
        NZ = Check1
    End If
    
End Function

Public Function AddBrackets(ObjName As String) As String
'this functions adds [] to object names that might need
'them because they have spaces in them
If InStr(ObjName, " ") > 0 And Mid(ObjName, 1, 1) <> "[" Then
    AddBrackets = "[" & ObjName & "]"
  Else
    AddBrackets = ObjName
  End If
End Function

Public Function StripBrackets(ObjName As String) As String
  'this function strips the [] off of data objects
  If Mid(ObjName, 1, 1) = "[" Then
    StripBrackets = Mid(ObjName, 2, Len(ObjName) - 2)
  Else
    StripBrackets = ObjName
  End If

End Function

Public Function StripFileName(rsFileName As String) As String
'this function strips the file name from a path\file string
  'On Error Resume Next
  Dim i As Integer

  For i = Len(rsFileName) To 1 Step -1
    If Mid(rsFileName, i, 1) = "\" Then
      Exit For
    End If
  Next

  StripFileName = Mid(rsFileName, 1, i - 1)

End Function

Public Function CheckNumberCHQ(ReadCheck As String, db As ADODB.Connection, Account As String, Optional CheckNumber As String) As String
Dim CheckNo As String
Dim rsBank As ADODB.Recordset
    Set rsBank = New ADODB.Recordset
      rsBank.Open "SELECT [BANK ACCT Next Check No] FROM [Bank Accounts] WHERE [BANK ACCT ID]='" & Account & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
      If rsBank.RecordCount = 0 Then
        MsgBox "Bank account is not valid!", , "Error"
        rsBank.Close
        Set rsBank = Nothing
        GoTo NoCheck
      End If
If CheckNumber = "" Then
    CheckNo = rsBank![BANK ACCT Next Check No] & ""
Else
    CheckNo = CheckNumber
End If

Select Case UCase(ReadCheck)
Case "READ"
    CheckNumber = CheckCheckNumber(Account, CheckNo, db, False)
    CheckNumberCHQ = CheckNumber
    rsBank![BANK ACCT Next Check No] = Str(Val(CheckNumber) + 1)
Case "CHECK"
    'samada CheckNumberCHQ= Found atau CheckNumberCHQ="Not Found"
    CheckNumberCHQ = CheckCheckNumber(Account, CheckNo, db, True)
    'CheckNumberCHQ = CheckNumber
    'If CheckNumberCHQ = "Not Found" Then
    '    rsBank![BANK ACCT Next Check No] = Str(Val(CheckNumber) + 1)
    'End If
Case "BACK"
    rsBank![BANK ACCT Next Check No] = Str(Val(CheckNumber) - 1)
End Select
    rsBank.Update
    rsBank.Close
    Set rsBank = Nothing
'013-3514255
Exit Function

NoCheck:
    CheckNumberCHQ = ""
    ShowStatus False
End Function

Public Function CheckCheckNumber(Account As String, CheckNumber As String, db As ADODB.Connection, Optional MustTrue As Boolean) As String
CheckAgain:
    If CheckDocument("SELECT [AP PAY Check No] FROM [AP Payment Header] WHERE [AP PAY Bank Account]='" & Account & "' AND [AP PAY Check No]='" & CheckNumber & "'", db, True) = True Then
        'cek masih tidak digunakan
        If MustTrue = True Then
            CheckCheckNumber = "Not Found"
            Exit Function
        Else
            CheckCheckNumber = Str(CheckNumber)
        End If
    Else
        'cek sudah digunakan
        If MustTrue = True Then
            CheckCheckNumber = "Found"
            Exit Function
        Else
            CheckNumber = Str(Val(CheckNumber) + 1)
            GoTo CheckAgain
        End If
    End If
End Function

Public Function CheckCreditLimit(CurrentRequest As Currency, txtFieldsCust As String, db As ADODB.Connection, Optional ShowCredit As Boolean) As Boolean
'repaired on 16/4/2000-razi '

  'Check Credit limit for this customer
  Dim Limit#
  Dim CurrentBalance#
  Dim Response%
  Limit# = LookRecord("[AR CUST Credit Limit]", "[AR Customer]", db, "[AR CUST Customer ID] = '" & txtFieldsCust & "'")
  If Limit# > 0 Then
    CurrentBalance# = LookRecord("[AR CUST Financial Period 1]", "[AR Customer]", db, "[AR CUST Customer ID] = '" & txtFieldsCust & "'")
    CurrentBalance# = CurrentBalance# + CurrentRequest
    If ShowCredit = False Then
        If CurrentBalance# > Limit# Then
          Response% = MsgBox("New balance will exceed " & txtFieldsCust & " credit limit!" & vbCr & vbCr & _
          "Credit Limit            : " & FormatCurr(CCur(Limit#)) & vbCr & _
          "Previous Balance : " & FormatCurr(CurrentBalance# - CurrentRequest) & vbCr & _
          "New Request       : " & FormatCurr(CurrentRequest) & vbCr & vbCr & _
          "Credit Balance      : " & FormatCurr(CCur(Limit# - CurrentBalance#)) & vbCr & vbCr & _
          "Would you like to Continue?", vbYesNo, "Information")
          If Response% = vbNo Then
            CheckCreditLimit = False
            Exit Function
          End If
        End If
    Else
        MsgBox txtFieldsCust & " Credit Limit Information." & vbCr & vbCr & _
        "Credit Limit            : " & FormatCurr(CCur(Limit#)) & vbCr & _
        "Previous Balance : " & FormatCurr(CurrentBalance# - CurrentRequest) & vbCr & _
        "New Request       : " & FormatCurr(CurrentRequest) & vbCr & _
        "Credit Balance      : " & FormatCurr(CCur(Limit# - CurrentBalance#)), vbInformation, "Credit Limit Information"
    End If
  Else
        MsgBox "There is no credit limit for " & txtFieldsCust
  End If
CheckCreditLimit = True
End Function

Public Function AcctBalance(RequestType As String, Accttype As String, db As ADODB.Connection, Optional ReqValue As Currency) As String
Dim AcctName As String

Select Case UCase(RequestType)
Case "BALANCE"
    AcctBalance = NZ(LookRecord("[GL COA Account Balance]", "[GL Chart Of Accounts]", db, "[GL COA Account No] = '" & Accttype & "'"))
    AcctBalance = FormatCurr(CCur(AcctBalance))
Case "NAME"
    AcctBalance = NZ(LookRecord("[GL COA Account Name]", "[GL Chart Of Accounts]", db, "[GL COA Account No] = '" & Accttype & "'"))
Case "DIFFERENCE"
    AcctBalance = NZ(LookRecord("[GL COA Account Balance]", "[GL Chart Of Accounts]", db, "[GL COA Account No] = '" & Accttype & "'"))
    If CCur(AcctBalance) - ReqValue < 0 Then
        AcctName = NZ(LookRecord("[GL COA Account Name]", "[GL Chart Of Accounts]", db, "[GL COA Account No] = '" & Accttype & "'"))
        MsgBox AcctName & "(" & Accttype & ") available balance (" & FormatCurr(CCur(AcctBalance)) & _
        ") is not enough to cover the request amount(" & FormatCurr(ReqValue) & ")"
        AcctBalance = "Not Enough"
    End If
End Select
End Function

Function PostGLWorkDetail(GLDate As Variant, GLNumber&, Optional db As ADODB.Connection)

  ''On Error Resume Next
  Dim AccountPost$
  Dim TranDate As Variant
  Dim DebitAmount@
  Dim CreditAmount@
  Dim Success%
  Dim Currentdb As Boolean
  Dim SQLstatement As String
  
  Currentdb = False
  'Dim db As ADODB.Connection
  If db Is Nothing Then
    Set db = New ADODB.Connection
    db.CursorLocation = adUseClient
    db.Open gblADOProvider
    Currentdb = True
  End If
  
  'Dim rsGLTransDetail As ADODB.Recordset
  'Set rsGLTransDetail = New ADODB.Recordset
  'rsGLTransDetail.Open "[GL Transaction Detail]", db, adOpenKeyset, adLockOptimistic, adCmdTable
  
  gLinesPosted% = 0

  ' exit if new GL # = 0
  If GLNumber& = 0 Then Exit Function

  ' exit if GL Date isn't a date
  If IsDate(GLDate) Then
  Else
    Exit Function
  End If

  ' post GL Work to GL Trans
  Dim DynaGL As ADODB.Recordset
  Dim GLWorkDebitAmount@
  Dim GLWorkCreditAmount@
  Dim GlDebitTotal@
  Dim GLCreditTotal@
  Dim NewGLNumber&
  
  GlDebitTotal@ = 0
  GLCreditTotal@ = 0

  Err = 0
  
  NewGLNumber& = 0
  
  ' create the dynaset summarized by Project & Account
  Dim GotCountFlag As Integer
FindAgain:
  Set DynaGL = New ADODB.Recordset
  DynaGL.Open "qrySumGLWorkDetail", db, adOpenKeyset, adLockReadOnly, adCmdTable
      
  If DynaGL.RecordCount = 0 Then
    ' no data
    GotCountFlag = MsgBox("Work is in Progress, Please Wait a Few second." & vbCr & "Stopping the process will cause an internal (Razi is still working on this)", vbYesNo, "Information")
    If GotCountFlag = vbYes Then
        DynaGL.Close
        GoTo FindAgain
    Else
        GoTo Skip_PostGLWorkDetail
    End If
  Else
    DynaGL.MoveFirst
  ''On Error GoTo PostGLWork_Error
    
    NewGLNumber& = GLNumber&

    Do While DynaGL.EOF = False
      If DynaGL("SumOfGW TRANSD Debit Amount") = DynaGL("SumOfGW TRANSD Credit Amount") Then
        ' They negate each other so don't write them
      Else
        If DynaGL("SumOfGW TRANSD Debit Amount") > 0 And DynaGL("SumOfGW TRANSD Credit Amount") > 0 Then

          ' Process debit portion
          GLWorkDebitAmount@ = DynaGL("SumOfGW TRANSD Debit Amount")
          GLWorkCreditAmount@ = 0
          '$
          ' accumulate debit & credit totals
          GlDebitTotal@ = GlDebitTotal@ + GLWorkDebitAmount@
          GLCreditTotal@ = GLCreditTotal@ + GLWorkCreditAmount@
    
          ' write the gl detail
          '-----------------------------------------------------
          'rsGLTransDetail.AddNew
          '  rsGLTransDetail("GL TRANSD Number") = NewGLNumber&
          '  rsGLTransDetail("GL TRANSD Account") = CStr(DynaGL("GW TRANSD Account"))
          '  rsGLTransDetail("GL TRANSD Debit Amount") = GLWorkDebitAmount@
          '  rsGLTransDetail("GL TRANSD Credit Amount") = GLWorkCreditAmount@
          '-----------------------------------------------------
            If IsNull(DynaGL("GW TRANSD Project")) Then
                SQLstatement = "INSERT INTO [GL Transaction Detail]"
                SQLstatement = SQLstatement & " ([GL TRANSD Number],[GL TRANSD Account],[GL TRANSD Debit Amount],[GL TRANSD Credit Amount])"
                SQLstatement = SQLstatement & " VALUES (" & NewGLNumber& & ",'" & CStr(DynaGL("GW TRANSD Account")) & "'," & GLWorkDebitAmount@ & "," & GLWorkCreditAmount@ & ")"
                db.Execute SQLstatement
            Else
                SQLstatement = "INSERT INTO [GL Transaction Detail]"
                SQLstatement = SQLstatement & " ([GL TRANSD Number],[GL TRANSD Account],[GL TRANSD Debit Amount],[GL TRANSD Credit Amount],[GL TRANSD Project])"
                SQLstatement = SQLstatement & " VALUES (" & NewGLNumber& & ",'" & CStr(DynaGL("GW TRANSD Account")) & "'," & GLWorkDebitAmount@ & "," & GLWorkCreditAmount@ & "," & CStr(DynaGL("GW TRANSD Project")) & ")"
                db.Execute SQLstatement
                'rsGLTransDetail("GL TRANSD Project") = CStr(DynaGL("GW TRANSD Project"))
            End If
          'rsGLTransDetail.Update
          '-----------------------------------------------------
  
          ' post to the Chart of Accounts
          AccountPost$ = DynaGL("GW TRANSD Account")
          TranDate = DateValue(GLDate)
          DebitAmount@ = GLWorkDebitAmount@
          CreditAmount@ = GLWorkCreditAmount@
          Success% = PostCOA(AccountPost$, TranDate, DebitAmount@, CreditAmount@, db)
          
          
          ' Process credit portion
          GLWorkDebitAmount@ = 0
          GLWorkCreditAmount@ = DynaGL("SumOfGW TRANSD Credit Amount")
          '
          ' accumulate debit & credit totals
          GlDebitTotal@ = GlDebitTotal@ + GLWorkDebitAmount@
          GLCreditTotal@ = GLCreditTotal@ + GLWorkCreditAmount@
    
          ' write the gl detail
          '------------------------------------------------------
          'rsGLTransDetail.AddNew
          '  rsGLTransDetail("GL TRANSD Number") = NewGLNumber&
          '  rsGLTransDetail("GL TRANSD Account") = DynaGL("GW TRANSD Account") & ""
          '  rsGLTransDetail("GL TRANSD Debit Amount") = GLWorkDebitAmount@
          '  rsGLTransDetail("GL TRANSD Credit Amount") = GLWorkCreditAmount@
          '  'rsGLTransDetail("GL TRANSD Project") = DynaGL("GW TRANSD Project") & ""
          '  If IsNull(DynaGL("GW TRANSD Project")) Then
          '  Else
          '      rsGLTransDetail("GL TRANSD Project") = CStr(DynaGL("GW TRANSD Project"))
          '  End If
            If IsNull(DynaGL("GW TRANSD Project")) Then
                SQLstatement = "INSERT INTO [GL Transaction Detail]"
                SQLstatement = SQLstatement & " ([GL TRANSD Number],[GL TRANSD Account],[GL TRANSD Debit Amount],[GL TRANSD Credit Amount])"
                SQLstatement = SQLstatement & " VALUES (" & NewGLNumber& & ",'" & DynaGL("GW TRANSD Account") & "" & "'," & GLWorkDebitAmount@ & "," & GLWorkCreditAmount@ & ")"
                db.Execute SQLstatement
            Else
                SQLstatement = "INSERT INTO [GL Transaction Detail]"
                SQLstatement = SQLstatement & " ([GL TRANSD Number],[GL TRANSD Account],[GL TRANSD Debit Amount],[GL TRANSD Credit Amount],[GL TRANSD Project])"
                SQLstatement = SQLstatement & " VALUES (" & NewGLNumber& & ",'" & DynaGL("GW TRANSD Account") & "" & "'," & GLWorkDebitAmount@ & "," & GLWorkCreditAmount@ & "," & CStr(DynaGL("GW TRANSD Project")) & ")"
                db.Execute SQLstatement
                'rsGLTransDetail("GL TRANSD Project") = CStr(DynaGL("GW TRANSD Project"))
            End If
          'rsGLTransDetail.Update
          '------------------------------------------------------
  
          ' post to the Chart of Accounts
          AccountPost$ = DynaGL("GW TRANSD Account")
          TranDate = DateValue(GLDate)
          DebitAmount@ = GLWorkDebitAmount@
          CreditAmount@ = GLWorkCreditAmount@
          Success% = PostCOA(AccountPost$, TranDate, DebitAmount@, CreditAmount@, db)
          If Success% = False Then GoTo PostGLWork_Error
          gLinesPosted% = gLinesPosted% + 1
          '$
        Else
          If DynaGL("SumOfGW TRANSD Debit Amount") > DynaGL("SumOfGW TRANSD Credit Amount") Then
            ' determine debit amount
            GLWorkDebitAmount@ = DynaGL("SumOfGW TRANSD Debit Amount") - DynaGL("SumOfGW TRANSD Credit Amount")
            GLWorkCreditAmount@ = 0
          Else
            ' determine credit amount
            GLWorkDebitAmount@ = 0
            GLWorkCreditAmount@ = DynaGL("SumOfGW TRANSD Credit Amount") - DynaGL("SumOfGW TRANSD Debit Amount")
          End If
  
          ' accumulate debit & credit totals
          GlDebitTotal@ = GlDebitTotal@ + GLWorkDebitAmount@
          GLCreditTotal@ = GLCreditTotal@ + GLWorkCreditAmount@
      
          ' write the gl detail
          '-------------------------------------------------------
          'rsGLTransDetail.AddNew
          '  rsGLTransDetail("GL TRANSD Number") = NewGLNumber&
          '  rsGLTransDetail("GL TRANSD Account") = CStr(DynaGL("GW TRANSD Account"))
          '  rsGLTransDetail("GL TRANSD Debit Amount") = GLWorkDebitAmount@
          '  rsGLTransDetail("GL TRANSD Credit Amount") = GLWorkCreditAmount@
          '  'rsGLTransDetail("GL TRANSD Project") = IIf(IsNull(DynaGL("GW TRANSD Project")), "", CStr(DynaGL("GW TRANSD Project")))
          '  If IsNull(DynaGL("GW TRANSD Project")) Then
          '  Else
          '      rsGLTransDetail("GL TRANSD Project") = CStr(DynaGL("GW TRANSD Project"))
          '  End If
            If IsNull(DynaGL("GW TRANSD Project")) Then
                SQLstatement = "INSERT INTO [GL Transaction Detail]"
                SQLstatement = SQLstatement & " ([GL TRANSD Number],[GL TRANSD Account],[GL TRANSD Debit Amount],[GL TRANSD Credit Amount])"
                SQLstatement = SQLstatement & " VALUES (" & NewGLNumber& & ",'" & CStr(DynaGL("GW TRANSD Account")) & "'," & GLWorkDebitAmount@ & "," & GLWorkCreditAmount@ & ")"
                db.Execute SQLstatement
            ElseIf CStr(DynaGL("GW TRANSD Project")) = "" Then
                SQLstatement = "INSERT INTO [GL Transaction Detail]"
                SQLstatement = SQLstatement & " ([GL TRANSD Number],[GL TRANSD Account],[GL TRANSD Debit Amount],[GL TRANSD Credit Amount])"
                SQLstatement = SQLstatement & " VALUES (" & NewGLNumber& & ",'" & CStr(DynaGL("GW TRANSD Account")) & "'," & GLWorkDebitAmount@ & "," & GLWorkCreditAmount@ & ")"
                db.Execute SQLstatement
            Else
                SQLstatement = "INSERT INTO [GL Transaction Detail]"
                SQLstatement = SQLstatement & " ([GL TRANSD Number],[GL TRANSD Account],[GL TRANSD Debit Amount],[GL TRANSD Credit Amount],[GL TRANSD Project])"
                SQLstatement = SQLstatement & " VALUES (" & NewGLNumber& & ",'" & DynaGL("GW TRANSD Account") & "" & "'," & GLWorkDebitAmount@ & "," & GLWorkCreditAmount@ & "," & CStr(DynaGL("GW TRANSD Project")) & ")"
                db.Execute SQLstatement
                'rsGLTransDetail("GL TRANSD Project") = CStr(DynaGL("GW TRANSD Project"))
            End If
          'rsGLTransDetail.Update
          '-------------------------------------------------------
    
          ' post to the Chart of Accounts
          AccountPost$ = DynaGL("GW TRANSD Account")
          TranDate = DateValue(GLDate)
          DebitAmount@ = GLWorkDebitAmount@
          CreditAmount@ = GLWorkCreditAmount@
          Success% = PostCOA(AccountPost$, TranDate, DebitAmount@, CreditAmount@, db)
          If Success% = False Then GoTo PostGLWork_Error
          gLinesPosted% = gLinesPosted% + 1
        End If
      End If  'Accounts are equal
      DynaGL.MoveNext
    Loop
  End If

    ' Update Gl Transaction Header
  Dim rsGLTrans As ADODB.Recordset
'again:
  Set rsGLTrans = New ADODB.Recordset
  rsGLTrans.Open "SELECT [GL Trans Number],[GL Trans Amount] FROM [GL Transaction] WHERE [GL Trans Number]=" & NewGLNumber& & "", db, adOpenKeyset, adLockOptimistic, adCmdText
  'MsgBox rsGLTrans.RecordCount
  'rsGLTrans.MoveLast
    If GlDebitTotal@ = GLCreditTotal@ Then
        rsGLTrans("GL Trans Amount") = GlDebitTotal@
        Else
        rsGLTrans("GL Trans Amount") = 0
    End If
  rsGLTrans.Update
  rsGLTrans.Close
  Set rsGLTrans = Nothing
 'GoTo again
  PostGLWorkDetail = True
  DynaGL.Close
  Set DynaGL = Nothing
  If Currentdb = True Then
    db.Close
    Set db = Nothing
  End If
  Exit Function

Skip_PostGLWorkDetail:

  DynaGL.Close
  Set DynaGL = Nothing
  If Currentdb = True Then
    db.Close
    Set db = Nothing
  End If
  ' End of post GL Work to GL Trans
  PostGLWorkDetail = False
  Exit Function

PostGLWork_Error:
  MsgBox "Error"
  DynaGL.Close
  PostGLWorkDetail = False
  Set DynaGL = Nothing
  If Currentdb = True Then
    db.Close
    Set db = Nothing
  End If

  Exit Function

End Function

Function PostCOA(Account$, TranDate As Variant, DebitAmount@, CreditAmount@, Optional db As ADODB.Connection) As Integer

  ''On Error Resume Next
  
  Dim AccountBalance@
  Dim BalanceType$
  Dim PeriodToPost%
  Dim PeriodClosed%
  Dim Post$
  
  If DebitAmount@ = 0 And CreditAmount@ = 0 Then
    PostCOA% = True
    Exit Function
  End If

Dim Currentdb As Boolean
  
  'Dim db As ADODB.Connection
  Currentdb = False
  If db Is Nothing Then
    Set db = New ADODB.Connection
    db.CursorLocation = adUseServer
    db.Open gblADOProvider
    Currentdb = True
  End If
  Call VerifyPeriod(TranDate, PeriodToPost%, PeriodClosed%, db)
  If PeriodToPost% = 0 Then
    Post$ = 1
  Else
    Post$ = Trim$(Str(PeriodToPost%))
  End If
  
  Dim rsGLCOA As ADODB.Recordset
  Set rsGLCOA = New ADODB.Recordset

  'rsGLCOA.Open "SELECT * FROM [GL Chart Of Accounts] WHERE [GL COA Account No]=" & Account$, db, adOpenStatic, adLockOptimistic, adCmdText '<<<---3 seconds
  rsGLCOA.Open "SELECT [GL COA Account No],[GL COA Balance Type]," & _
  "[GL COA CY Period " & Post$ & " Amt],[GL COA Account Balance],[GL COA CY Beginning Amt] " & _
  "FROM [GL Chart Of Accounts] WHERE [GL COA Account No]='" & Account$ & "'", db, adOpenKeyset, adLockOptimistic, adCmdText '<<<---3 seconds
  
  Post$ = Trim$(Str(PeriodToPost%))

  'rsGLCOA.Index = "PrimaryKey"
  'rsGLCOA.Find "[GL COA Account No]=" & Account$
  If rsGLCOA.RecordCount = 0 Then
    PostCOA = False
    Exit Function
  Else
  rsGLCOA.MoveFirst
  
    BalanceType$ = IIf(IsNull(rsGLCOA("GL COA Balance Type")), "Debit", rsGLCOA("GL COA Balance Type"))
    
    'Call VerifyPeriod(TranDate, PeriodToPost%, PeriodClosed%, db)
    'Post$ = Trim$(Str(PeriodToPost%))

    ''On Error GoTo PostCOA_Error

    If PeriodToPost% > 0 Then
      ' Post to Period
        AccountBalance@ = IIf(IsNull(rsGLCOA("GL COA CY Period " & Post$ & " Amt")), 0, rsGLCOA("GL COA CY Period " & Post$ & " Amt"))
        If DebitAmount@ <> 0 Then
          rsGLCOA("GL COA Account Balance") = rsGLCOA("GL COA Account Balance") + DebitAmount@
          rsGLCOA("GL COA CY Period " & Post$ & " Amt") = AccountBalance@ + DebitAmount@
        End If
        If CreditAmount@ <> 0 Then
          rsGLCOA("GL COA Account Balance") = rsGLCOA("GL COA Account Balance") - CreditAmount@
          rsGLCOA("GL COA CY Period " & Post$ & " Amt") = AccountBalance@ - CreditAmount@
        End If
      rsGLCOA.Update
      ' end of post to period
    Else
      'Post to Beginning Balance Account
        AccountBalance@ = IIf(IsNull(rsGLCOA("GL COA CY Beginning Amt")), 0, rsGLCOA("GL COA CY Beginning Amt"))
        If DebitAmount@ <> 0 Then
          rsGLCOA("GL COA Account Balance") = rsGLCOA("GL COA Account Balance") + DebitAmount@
          rsGLCOA("GL COA CY Beginning Amt") = AccountBalance@ + DebitAmount@
        End If
        If CreditAmount@ <> 0 Then
          rsGLCOA("GL COA Account Balance") = rsGLCOA("GL COA Account Balance") - CreditAmount@
          rsGLCOA("GL COA CY Beginning Amt") = AccountBalance@ - CreditAmount@
        End If
      rsGLCOA.Update
    End If
  End If
  
  MatchBank db, Account$, CDbl(DebitAmount@), CDbl(CreditAmount@)
  
  PostCOA = True
  rsGLCOA.Close
  Set rsGLCOA = Nothing
   If Currentdb = True Then
    db.Close
    Set db = Nothing
   End If
  
  Exit Function

PostCOA_Error:
  
  PostCOA = False
  rsGLCOA.Close
  Set rsGLCOA = Nothing
   If Currentdb = True Then
    db.Close
    Set db = Nothing
   End If

  Exit Function

End Function

Public Function FormatDate(RequestDate As Date, Optional DateType As String) As String
Select Case DateType
Case "MonthYear"
    FormatDate = Format(RequestDate, "mmmm yyyy")
Case "Long Date"
    FormatDate = Format(RequestDate, "Long Date")
Case "General Date"
    FormatDate = Format(RequestDate, "General Date")
Case "Medium Date"
    FormatDate = Format(RequestDate, "Medium Date")
Case "Short Date"
    FormatDate = Format(RequestDate, "Short Date")
Case Else
    FormatDate = Format(RequestDate, "mm/dd/yyyy")
End Select
End Function

Public Function FormatCurr(RequestCurr As Currency, Optional AcctStandard As Boolean) As String
If AcctStandard = True Then
    FormatCurr = Format(RequestCurr, "$###,###,###,##0.00")
Else
    FormatCurr = Format(RequestCurr, "$###,###,###,##0.00;($###,###,###,##0.00)")
End If
End Function

Public Function ValidatePower(DocumentNo As String, TransType As String, TransProcess As String, db As ADODB.Connection) As Boolean
Dim Power As String
Dim SQLstatement As String

  Power = LookRecord("[EMP Custom 1]", "[EMP Employees]", db, "[EMP ID] = '" & AppLoginName & "'")
  If CInt(Power) = 1 Then
      SQLstatement = "INSERT INTO [Transaction]"
      SQLstatement = SQLstatement & " ([Transaction ID],[Transaction Approved By],[Transaction Approved On],[Transaction Process],[Transaction Type])"
      SQLstatement = SQLstatement & " VALUES ('" & DocumentNo & "','" & AppLoginName & "',#" & Now & "#,'" & TransProcess & "','" & TransType & "')"
      db.Execute SQLstatement
      'Debug.Print SQLstatement
    ValidatePower = True
  Else
    MsgBox "You have no authority to approved this transaction", vbInformation, "Information"
    ValidatePower = False
  End If

End Function
