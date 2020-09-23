Attribute VB_Name = "Primary_SUB"
Option Explicit

Public Sub ConvertToCurr(strValue As TextBox)
      strValue.Text = FormatCurr(strValue.Text)
      If Left(strValue.Text, 1) <> "$" Then
         MsgBox "Please enter a number", vbInformation, "Error"
         strValue.Text = "$0.00"
      End If
End Sub

Public Sub ProcessDoneMusic(ProcessType As String)
Select Case ProcessType
Case "Done"
    Beep
End Select
End Sub

Public Sub ComboInit(ComboName As ComboBox, labelName As label, SQLstatement As String, Optional db As ADODB.Connection)
'Dim db As ADODB.Connection
Dim cbRS As ADODB.Recordset
Dim i As Integer
Dim LoadDB As Boolean

LoadDB = False
If db Is Nothing Then
  Set db = New ADODB.Connection
  db.CursorLocation = adUseClient
  db.Open gblADOProvider
  LoadDB = True
End If

  Dim TempCbData As String
  TempCbData = ComboName.Text
  Set cbRS = New ADODB.Recordset
  i = 0
  With cbRS
  Select Case labelName
  Case "Payments Methods"
        .Open SQLstatement & " WHERE [Payment Terms]='" & fMainForm.ActiveForm.cbPurchase(5).Text & "'", db, adOpenKeyset, adLockReadOnly, adCmdText
  Case Else
        .Open SQLstatement, db, adOpenKeyset, adLockReadOnly, adCmdText
  End Select
    ComboName.Clear
    If .RecordCount = 0 Then
       'trap this error
       MsgBox "There is no data for " & labelName.Caption
    Else
        
        Do While Not .EOF
          ComboName.List(i) = cbRS(0).Value
          i = i + 1
           .MoveNext
        Loop
    End If
  .Close
  End With
    ComboName.Text = TempCbData
    CheckCombo ComboName

Set cbRS = Nothing
If LoadDB = True Then
    db.Close
    Set db = Nothing
End If
End Sub

Public Sub EndLoad(db As ADODB.Connection, ADOprimaryrs As ADODB.Recordset, rsField As String, Optional whichOne As Integer)
On Error GoTo EndLoadErr
    Dim chklst As ADODB.Recordset
    Set chklst = New ADODB.Recordset
    
    chklst.Open "SELECT [" & rsField & "] from [SYS Setup]", db, adOpenKeyset, adLockOptimistic, adCmdText
    If chklst.RecordCount = 0 Then
        chklst.AddNew
    End If
    ADOprimaryrs.CancelUpdate
    If ADOprimaryrs.RecordCount > 0 Then
        chklst.Fields("" & rsField & "").Value = True
    Else
        chklst.Fields("" & rsField & "").Value = False
    End If
    chklst.Update
    chklst.Requery
    chklst.Close
    Set chklst = Nothing
    
    'If whichOne <> 1 Then
        'check the checklist
        IntroCheckList
    'End If
EndLoadErr:
End Sub

Public Sub IntroCheckList(Optional StatusList As Integer)
On Error GoTo IntroErr

Dim cnn As ADODB.Connection
Dim chklst As ADODB.Recordset

Dim AllCheck As Boolean
Dim FieldCount As Integer

  'get connection name
  DbConnectionString gblApplicationConnectString
    
  Set cnn = New ADODB.Connection
  cnn.CursorLocation = adUseServer
  cnn.Open gblADOProvider
    
    AllCheck = True
    
    Set chklst = New ADODB.Recordset
    
    chklst.Open "SELECT * from [SYS Setup]", cnn, adOpenKeyset, adLockOptimistic, adCmdText
    If chklst.RecordCount > 0 Then
        FieldCount = chklst.Fields.count
        Dim i As Integer
            For i = 0 To FieldCount - 1
                If chklst(i) = False Then
                    Select Case i
                    Case 10, 15, 16
                    Case Else
                        AllCheck = False
                        Exit For
                    End Select
                End If
            Next
    Else
        AllCheck = False
    End If
    If AllCheck = False Then
        'frm_SYS_Setup_Checklist.Show
        If StatusList = 1 Then
            MsgBox "Setup Module is not completed in " & chklst(i).Name & " , you have to complete the Module before you could continue to the other module." & _
            "Use the checklist to complete it.", vbInformation, "Setup Module Not Completed"
        End If
        fMainForm.MenuStatus True, True, False
    Else
       fMainForm.MenuStatus True, True, True
    End If
    chklst.Close
    Set chklst = Nothing
    cnn.Close
    Set cnn = Nothing

Exit Sub
IntroErr:
    MsgBox "An error occured while attempting to read" & vbCr & gblApplicationConnectString & vbCr & "Either the database is corrupted or someone has altered one of the Table" & vbCr & _
    "Please load another Database", vbCritical, "Information"
    fMainForm.MenuStatus False, False, False
    Set chklst = Nothing
    cnn.Close
    Set cnn = Nothing

End Sub

Public Sub RefreshButton(ADOprimaryrs As ADODB.Recordset, grddatagrid As DataGrid)
ShowStatus True
Dim mvBookMark As Variant

'This is only needed for multi user apps
'On Error GoTo RefreshErr

If ADOprimaryrs.RecordCount > 0 Then
  
  Set grddatagrid.DataSource = Nothing
  If ADOprimaryrs.RecordCount > 1 Then
        If ADOprimaryrs.EditMode <> 0 Then
            ADOprimaryrs.UpdateBatch adAffectAll
        End If
        
        mvBookMark = ADOprimaryrs.Bookmark
          If ADOprimaryrs.EditMode <> 0 Then
            ADOprimaryrs.CancelUpdate
            ADOprimaryrs.Requery
          End If
        ADOprimaryrs.Bookmark = mvBookMark
  Else
        ADOprimaryrs.UpdateBatch adAffectAll
        ADOprimaryrs.Requery
  End If
  
  Set grddatagrid.DataSource = ADOprimaryrs("ChildCMD").UnderlyingValue
ShowStatus False
End If
  Exit Sub
RefreshErr:
  MsgBox Err.Description
  Resume Next
End Sub

Public Sub UpdateButton(ADOprimaryrs As ADODB.Recordset, mbAddNewFlag As Boolean)
'On Error GoTo UpdateErr

If ADOprimaryrs.RecordCount > 0 Then

  ADOprimaryrs.Update

  'If mbAddNewFlag Then
  '  ADOprimaryRS.MoveLast              'move to the new record
  'End If
  mbAddNewFlag = False
End If
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Public Sub DueDateDay(db As ADODB.Connection, cbPurchase As ComboBox, TxtDate As TextBox, txtTarget As TextBox, Optional ShowMessage As Boolean)
On Error GoTo DueDate_Error
  
  ' setup the due date based on the customer terms
  
  Dim rsPaymentTerms As ADODB.Recordset
  Set rsPaymentTerms = New ADODB.Recordset
  rsPaymentTerms.Open ("[LIST Payment Terms]"), db, adOpenStatic, adLockOptimistic, adCmdTable
  
  Dim PaymentTermsSales$
  Dim DiscountDays%
  Dim DiscountDate As Variant
  Dim DueDate As Variant

  PaymentTermsSales$ = IIf(IsNull(cbPurchase.Text), "", cbPurchase.Text)

  If PaymentTermsSales$ = "" Then
    If ShowMessage = False Then MsgBox "Payment terms is Empty. Default value will be used", vbCritical, "Error"
    DueDate = DateAdd("d", 30, TxtDate.Text)
  Else
    rsPaymentTerms.MoveFirst
    rsPaymentTerms.Find "[LIST PAY Description]='" & PaymentTermsSales$ & "'"
    If rsPaymentTerms.EOF = False Then
      DueDate = DateAdd("d", rsPaymentTerms("List Pay Net Days"), TxtDate.Text)
      DiscountDays% = IIf(IsNull(rsPaymentTerms("List Pay Discount Days")), 0, rsPaymentTerms("List Pay Discount Days"))
      If DiscountDays% > 0 Then
        DiscountDate = DateAdd("d", rsPaymentTerms("List Pay Discount Days"), TxtDate.Text)
      Else
        DiscountDate = TxtDate.Text
      End If
    Else
      DueDate = DateAdd("d", 30, TxtDate.Text)
      DiscountDate = TxtDate.Text
    End If
  End If

  'On Error Resume Next
   txtTarget = FormatDate(CDate(DueDate))
  'If lblfields(6).Caption = "Due Date" Then
  '   ![AR ORDER Other Date] = Format(DueDate, "Short Date")
  'End If
  Exit Sub
DueDate_Error:
  Call ErrorLog("Order Transactions", "DueDateDay", Now, Err.Number, Err.Description, True, db)
  Resume Next
End Sub

Public Sub NewRowForDataGrid(ADOprimaryrs As ADODB.Recordset, grddatagrid As DataGrid, FieldName As String, FieldData As String)
'On Error GoTo NewErr
Dim mvBookMark As Variant
          
     With ADOprimaryrs
        ADOprimaryrs("" & FieldName & "") = FieldData
        ADOprimaryrs.Update
        If .BOF = False Or .EOF = False Then
           mvBookMark = .Bookmark
        End If
        Set grddatagrid.DataSource = Nothing
            ADOprimaryrs.Requery
        Set grddatagrid.DataSource = ADOprimaryrs("ChildCMD").UnderlyingValue
        If mvBookMark > 0 Then
           ADOprimaryrs.Bookmark = mvBookMark '+ 1
        Else
           ADOprimaryrs.MoveFirst
        End If
    End With
    grdOnAddNew = False
Exit Sub
NewErr:
  MsgBox Err.Description
End Sub

Public Sub GetTextColor(TargetForm As Form)
On Error GoTo NewErr
ShowStatus True

Dim TemptConnection As String
Dim TempgblADOProvider As String

Dim ADOprimaryrs As ADODB.Recordset
Dim dbTemp As ADODB.Connection
Dim SQLstatement As String
      
      Set dbTemp = New ADODB.Connection
      dbTemp.CursorLocation = adUseClient
      dbTemp.Open gblBasicADOProvider
      
      SQLstatement = "SELECT * FROM Properties"
      Set ADOprimaryrs = New ADODB.Recordset
      ADOprimaryrs.Open SQLstatement, dbTemp, adOpenKeyset, adLockOptimistic, adCmdText
 
 With ADOprimaryrs
 Dim Ctrl As Control
 For Each Ctrl In TargetForm.Controls
   'MsgBox Ctrl.Name
   'If TypeOf Ctrl Is TextBox Or TypeOf Ctrl Is CheckBox Or TypeOf Ctrl Is ComboBox Then
   If TypeOf Ctrl Is TextBox Then
    If Ctrl.Enabled = False Then
        .MoveFirst
        .Find "Type='Disabled Input Box'"
    ElseIf Ctrl.Locked = True Then
        .MoveFirst
        .Find "Type='Locked Input Box'"
    Else
        .MoveFirst
        .Find "Type='Standard Input Box'"
    End If
    
        Ctrl.BackColor = ADOprimaryrs("BackColor").Value
        Ctrl.Appearance = ADOprimaryrs("Appearance").Value
        Ctrl.BorderStyle = ADOprimaryrs("BorderStyle").Value
        Ctrl.Font = ADOprimaryrs("Font").Value
        Ctrl.ForeColor = ADOprimaryrs("ForeColor").Value
   End If
   
   'For labels
   If TypeOf Ctrl Is label Then
    If UCase(Ctrl.Name) = "LBLFIELDS" Then
        .MoveFirst
        .Find "Type='Above Label'"
        Ctrl.Height = 285
        Ctrl.BackStyle = 1
    ElseIf UCase(Ctrl.Name) = "LBLLABELS" Then
        .MoveFirst
        .Find "Type='Side Label'"
        Ctrl.BackStyle = 1
        Ctrl.Height = 285
    Else
        GoTo JumpLoop
    End If
        Ctrl.Appearance = ADOprimaryrs("Appearance").Value
        Ctrl.BorderStyle = ADOprimaryrs("BorderStyle").Value
        Ctrl.Font = ADOprimaryrs("Font").Value
        Ctrl.ForeColor = ADOprimaryrs("ForeColor").Value
        Ctrl.BackColor = ADOprimaryrs("BackColor").Value
   End If
   
   'for datagrid
   If TypeOf Ctrl Is DataGrid Then
        .MoveFirst
        .Find "Type='Table'"
        Ctrl.BackColor = ADOprimaryrs("BackColor").Value
        Ctrl.Appearance = ADOprimaryrs("Appearance").Value
        Ctrl.BorderStyle = ADOprimaryrs("BorderStyle").Value
        'Ctrl.Font = ADOprimaryrs("Font").Value
        Ctrl.ForeColor = ADOprimaryrs("ForeColor").Value
   End If
JumpLoop:
 Next
         .MoveFirst
         .Find "Type='Interface'"
         TargetForm.BackColor = ADOprimaryrs("BackColor").Value
  Dim CtrlInter As Control
  For Each CtrlInter In TargetForm.Controls
    If TypeOf CtrlInter Is Frame Then
        CtrlInter.Appearance = ADOprimaryrs("Appearance").Value
        CtrlInter.BorderStyle = ADOprimaryrs("BorderStyle").Value
        CtrlInter.Font = ADOprimaryrs("Font").Value
        CtrlInter.ForeColor = ADOprimaryrs("ForeColor").Value
        CtrlInter.BackColor = ADOprimaryrs("BackColor").Value
    ElseIf TypeOf CtrlInter Is PictureBox Then
        'Picture2.Appearance = ADOprimaryrs("Appearance").Value
        'Picture2.BorderStyle = ADOprimaryrs("BorderStyle").Value
        CtrlInter.Font = ADOprimaryrs("Font").Value
        CtrlInter.ForeColor = ADOprimaryrs("ForeColor").Value
        CtrlInter.BackColor = ADOprimaryrs("BackColor").Value
    ElseIf TypeOf CtrlInter Is CheckBox Or TypeOf CtrlInter Is OptionButton Then
        CtrlInter.Appearance = ADOprimaryrs("Appearance").Value
        'Option1.BorderStyle = ADOprimaryrs("BorderStyle").Value
        CtrlInter.Font = ADOprimaryrs("Font").Value
        CtrlInter.ForeColor = ADOprimaryrs("ForeColor").Value
        CtrlInter.BackColor = ADOprimaryrs("BackColor").Value
    End If
  Next
 
 End With
    ADOprimaryrs.Close
    Set ADOprimaryrs = Nothing
    dbTemp.Close
    Set dbTemp = Nothing
ShowStatus False
Exit Sub
NewErr:
  MsgBox "Missing database, someone has deleted the database. Please contact the suppllier", vbCritical, "Error"
  
End Sub


Public Sub GetStartUp(WhichToolBar As Integer, ToolBarPic As String)
On Error GoTo NewErr

Dim TemptConnection As String
Dim TempgblADOProvider As String

Dim ADOprimaryrs As ADODB.Recordset
Dim dbTemp As ADODB.Connection

Dim SQLstatement As String

      Set dbTemp = New ADODB.Connection
      dbTemp.CursorLocation = adUseClient
      dbTemp.Open gblBasicADOProvider
      
      SQLstatement = "SELECT * FROM Properties WHERE [Type]='Startup'"
      Set ADOprimaryrs = New ADODB.Recordset
      ADOprimaryrs.Open SQLstatement, dbTemp, adOpenKeyset, adLockOptimistic, adCmdText
 
 With ADOprimaryrs
    WhichToolBar = ![Appearance]
    ToolBarPic = ![BackColor]
 End With
 
    ADOprimaryrs.Close
    Set ADOprimaryrs = Nothing
    dbTemp.Close
    Set dbTemp = Nothing

Exit Sub
NewErr:
  MsgBox "Missing database, someone has deleted the database. Please contact the suppllier", vbCritical, "Error"
  
End Sub

Public Sub TxtGotFocus(txtFields As TextBox)
With txtFields
  If .Enabled <> False Then
     .SelStart = 0
     .SelLength = Len(.Text)
  End If
    'SendKeys ("{End}")
End With
End Sub

Sub VerifyPeriod(TranDate As Variant, PeriodToPost%, PeriodClosed%, Optional db As ADODB.Connection)
Dim Currentdb As Boolean

  ReDim Period(1 To 14) As Variant
  Dim HighestPeriod%
  Dim X%
  
  'On Error Resume Next

  'Dim db As ADODB.Connection
  Currentdb = False
  If db Is Nothing Then
    Set db = New ADODB.Connection
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
  Set rsCompany = New ADODB.Recordset
  rsCompany.Open SQLstatement, db, adOpenKeyset, adLockOptimistic, adCmdText '<<<---3 seconds

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
  If PeriodToPost% <= -1 Then PeriodToPost% = HighestPeriod%
    ' End of Test Period to Post
  'If PeriodToPost% = 0 Then
  '  MsgBox "The input date [" & Format(TranDate, "mm/dd/yyyy") & "] value is before the FISCAL START DATE," & vbCr & " it will be posted to the beginning of the period [" & Period(14) & "]", vbCritical, "Information"
  '  PeriodToPost% = 1
  'End If

  'Determine if period to post is closed
  If PeriodToPost% >= 14 Then
      MsgBox "The input date [" & FormatDate(CDate(TranDate)) & "] value is exceeded the FISCAL END DATE," & vbCr & " it will be posted to the end of the period [" & Period(14) & "]", vbCritical, "Information"
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

Public Sub SearchRECORD(ADOprimaryrs As ADODB.Recordset, grddatagrid As DataGrid, txt As String, Labels As String, WhichField As String, DefaultField As String)
On Error GoTo NOTFOUND
Dim mvBookMark As Variant
If txt = "" Then Exit Sub
    grddatagrid.SetFocus
    
    If Not (ADOprimaryrs.BOF Or ADOprimaryrs.EOF) Then
      mvBookMark = ADOprimaryrs.Bookmark
    End If
    
    ADOprimaryrs.MoveFirst
    
    If WhichField = "" Then
        WhichField = DefaultField
    End If
    If ADOprimaryrs("" & WhichField & "").Type = 202 Then
        ADOprimaryrs.Find "[" & WhichField & "]='" & txt & "'"
    Else
        ADOprimaryrs.Find "[" & WhichField & "]=" & txt
    End If
    
    If ADOprimaryrs.EOF Then
NOTFOUND:
        MsgBox Labels & " " & txt & " is not existed.", vbInformation, "Information"
        If mvBookMark > 0 Then
          ADOprimaryrs.Bookmark = mvBookMark
        Else
          ADOprimaryrs.MoveFirst
        End If
    End If
    SendKeys ("{LEFT}")
End Sub
Public Sub CustomerData(FormName As Form, db As ADODB.Connection, CustID As String, typeRequest As Boolean)
ShowStatus True
Dim ADOcustRS As ADODB.Recordset
Dim SQLstatement As String

Set ADOcustRS = New ADODB.Recordset
SQLstatement = "SELECT [AR CUST Web Page],[AR EMail Address],[AR CUST Payment Terms]," & _
"[AR CUST Tax Group],[AR CUST Discount %] FROM [AR Customer] WHERE [AR CUST Customer ID]='" & _
CustID & "'"

ADOcustRS.Open SQLstatement, db, adOpenKeyset, adLockReadOnly, adCmdText
If ADOcustRS.RecordCount = 0 Then Exit Sub
With ADOcustRS
  If IsNull(![AR CUST Web Page]) Or Trim(![AR CUST Web Page]) = "" Then
    FormName.lblweb.Visible = False
  Else
    FormName.lblweb.ToolTipText = ![AR CUST Web Page]
    FormName.lblweb.Visible = True
  End If
  If IsNull(![AR EMail Address]) Or Trim(![AR EMail Address]) = "" Then
    FormName.lblmail.Visible = False
  Else
    FormName.lblmail.ToolTipText = ![AR EMail Address]
    FormName.lblmail.Visible = True
  End If
  
    If typeRequest = True Then
        If IsNull(![AR CUST Payment Terms]) Or Trim(![AR CUST Payment Terms]) = "" Then
        Else
            FormName.cbPurchase(5).Text = ![AR CUST Payment Terms]
        End If
        If IsNull(![AR CUST Tax Group]) Or Trim(![AR CUST Tax Group]) = "" Then
        Else
            FormName.cbPurchase(1).Text = ![AR CUST Tax Group]
        End If
        If IsNull(![AR CUST Discount %]) Then
        Else
            FormName.txtFields(28).Text = Format(![AR CUST Discount %], "00.00")
        End If
    End If
End With
ADOcustRS.Close
Set ADOcustRS = Nothing
ShowStatus False
End Sub

Public Sub VendorID(SQLstatement As String, db As ADODB.Connection, frm As Form)
Dim ADOVendorRS As ADODB.Recordset
Dim Response As Integer
Dim i As Integer

Set ADOVendorRS = New ADODB.Recordset
With ADOVendorRS
    .Open SQLstatement, db, adOpenKeyset, adLockReadOnly, adCmdText
    If .RecordCount > 0 Then
        For i = 0 To frm.txtFieldsVendor.UBound
            frm.txtFieldsVendor(i).Text = .Fields(i).Value
        Next
    Else
        Response = MsgBox("New Vendor ID, would you like to add it to the database?", vbYesNo, "Information")
        If Response = vbYes Then
            frm_AP_Vendor.CallByUserVendor frm.txtFieldsVendor(0).Text
            frm_AP_Vendor.ZOrder 0
            MsgBox "New data have been transferred to VENDIR SETUP, but you have to add more data and save it.", vbInformation, "Information"
        End If
        For i = 0 To frm.txtFieldsVendor.UBound
            frm.txtFieldsVendor(i).Text = ""
        Next
    End If
End With
ADOVendorRS.Close
Set ADOVendorRS = Nothing
End Sub

Public Sub CustomerID(SQLstatement As String, db As ADODB.Connection, frm As Form)
Dim ADOcustidRS As ADODB.Recordset
Dim ADOshiptoRS As ADODB.Recordset

Dim i As Integer

Set ADOcustidRS = New ADODB.Recordset
With ADOcustidRS
    .Open SQLstatement, db, adOpenKeyset, adLockReadOnly, adCmdText
    If .RecordCount > 0 Then
      If .Fields("AR CUST SalesPerson") <> AppLoginName Then
        Dim Response As Integer
            Response = MsgBox("This account is belong to " & .Fields("AR CUST SalesPerson") & _
            "." & vbCr & "Would you like to continue?", vbYesNo, "Information")
            If Response = vbYes Then
                For i = 0 To frm.txtFieldsCust.UBound
                    frm.txtFieldsCust(i).Text = .Fields(i).Value
                Next
                For i = 0 To frm.txtFieldsShip.UBound
                    frm.txtFieldsShip(i).Text = ""
                Next
            shipToID db, frm
            Else
                For i = 0 To frm.txtFieldsCust.UBound
                    frm.txtFieldsCust(i).Text = ""
                Next
                For i = 0 To frm.txtFieldsShip.UBound
                    frm.txtFieldsShip(i).Text = ""
                Next
            End If
       Else
            For i = 0 To frm.txtFieldsCust.UBound
                frm.txtFieldsCust(i).Text = .Fields(i).Value
            Next
            For i = 0 To frm.txtFieldsShip.UBound
                frm.txtFieldsShip(i).Text = ""
            Next
            shipToID db, frm
       End If
    Else
    Response = MsgBox("New Customer ID , would you like to add it to the database?", vbYesNo, "Information")
    If Response = vbYes Then
            frm_AR_Customer.CallByUserCust frm.txtFieldsCust(0).Text
            frm_AR_Customer.ZOrder 0
            MsgBox "New data have been transferred to CUSTOMER SETUP, but you have to add more data and save it.", vbInformation, "Information"
    End If
        For i = 0 To frm.txtFieldsCust.UBound
            frm.txtFieldsCust(i).Text = ""
        Next
        For i = 0 To frm.txtFieldsShip.UBound
            frm.txtFieldsShip(i).Text = ""
        Next
    End If
    
    .Close
End With

Set ADOcustidRS = Nothing
End Sub

Public Sub shipToID(db As ADODB.Connection, frm As Form)
Dim ADOshiptoRS As ADODB.Recordset
Dim i As Integer
Dim Response As Integer

Set ADOshiptoRS = New ADODB.Recordset
If frm.txtFieldsCust(0) = "" Then Exit Sub
If frm.txtFieldsShip(0).Text = "" Then
    ADOshiptoRS.Open "Select [AR SHIP ID],[AR SHIP Name],[AR SHIP Address 1]," & _
        "[AR SHIP Address 2],[AR SHIP City],[AR SHIP State],[AR SHIP Postal],[AR SHIP Country],[AR SHIP Phone],[AR SHIP Fax],[AR SHIP Default] From " & _
        "[AR SHIP to] WHERE [AR SHIP Customer ID]='" & frm.txtFieldsCust(0) & "' ", db, adOpenKeyset, adLockReadOnly, adCmdText
Else
    'Debug.Print "Select [AR SHIP ID],[AR SHIP Name],[AR SHIP Address 1]," & _
        "[AR SHIP Address 2],[AR SHIP City],[AR SHIP State],[AR SHIP Postal],[AR SHIP Country],[AR SHIP Phone],[AR SHIP Fax],[AR SHIP Default] From " & _
        "[AR SHIP to] WHERE [AR SHIP Customer ID]='" & frm.txtFieldsCust(0) & "' " & _
        "AND [AR SHIP Name]='" & frm.txtFieldsShip(0).Text & "'"
    ADOshiptoRS.Open "Select [AR SHIP ID],[AR SHIP Name],[AR SHIP Address 1]," & _
        "[AR SHIP Address 2],[AR SHIP City],[AR SHIP State],[AR SHIP Postal],[AR SHIP Country],[AR SHIP Phone],[AR SHIP Fax],[AR SHIP Default] From " & _
        "[AR SHIP to] WHERE [AR SHIP Customer ID]='" & frm.txtFieldsCust(0) & "' " & _
        "AND [AR SHIP ID]='" & frm.txtFieldsShip(0).Text & "'", db, adOpenKeyset, adLockReadOnly, adCmdText
End If
    'MsgBox frm.txtFieldsCust(0) & "      " & frm.txtFieldsShip(0).Text
    If ADOshiptoRS.RecordCount > 0 Then
        If ADOshiptoRS.RecordCount > 1 Then
            ADOshiptoRS.Find "[AR SHIP Default]=True"
            If ADOshiptoRS.EOF Then
                ADOshiptoRS.MoveFirst
            End If
        End If
        For i = 0 To frm.txtFieldsShip.UBound
            If IsNull(ADOshiptoRS.Fields(i).Value) Then
            Else
                frm.txtFieldsShip(i).Text = ADOshiptoRS.Fields(i).Value
            End If
        Next
    Else
      Response = MsgBox("This is a new input for SHIP TO SETUP , would you like to add it to the database?", vbYesNo, "Information")
      If Response = vbYes Then
        frm_AR_Cust_Ship_To.CallByUserShip frm.txtFieldsCust(0).Text
        MsgBox "New data have been transferred to SHIP TO SETUP, but you have to add more data and save it.", vbInformation, "Information"
      End If
        For i = 0 To frm.txtFieldsShip.UBound
            frm.txtFieldsShip(i).Text = ""
        Next
    End If
ADOshiptoRS.Close
Set ADOshiptoRS = Nothing
End Sub

Public Sub GetWEBMAILvendor(VendorID As String, db As ADODB.Connection, frm As Form)
Dim ADOmailWebRS As ADODB.Recordset

Set ADOmailWebRS = New ADODB.Recordset
ADOmailWebRS.Open "SELECT [AP VEN Custom 1],[AP VEN Custom 2] FROM [AP Vendor] WHERE [AP VEN ID]='" & VendorID & "'", db, adOpenKeyset, adLockReadOnly, adCmdText
With ADOmailWebRS
End With
    If IsNull(ADOmailWebRS![AP VEN Custom 1]) Or Trim(ADOmailWebRS![AP VEN Custom 1]) = "" Then
        frm.lblweb.Visible = False
    Else
        frm.lblweb.ToolTipText = ADOmailWebRS![AP VEN Custom 1]
        frm.lblweb.Visible = True
    End If
    
    If IsNull(ADOmailWebRS![AP VEN Custom 2]) Or Trim(ADOmailWebRS![AP VEN Custom 2]) = "" Then
        frm.lblmail.Visible = False
    Else
        frm.lblmail.ToolTipText = ADOmailWebRS![AP VEN Custom 2]
        frm.lblmail.Visible = True
    End If

ADOmailWebRS.Close
Set ADOmailWebRS = Nothing
End Sub

Public Sub ShowStatus(ProgStatus As Boolean)
On Error Resume Next
If ProgStatus = True Then
    Screen.MousePointer = vbHourglass
    fMainForm.picIcon.Picture = fMainForm.imgIcon.ListImages("off").Picture
Else
    Screen.MousePointer = vbNormal
    fMainForm.picIcon.Picture = fMainForm.imgIcon.ListImages("on").Picture
End If
End Sub

Public Sub MatchBank(db As ADODB.Connection, AccountNo As String, DebitAmount As Double, CreditAmount As Double)
Dim ADObankRS As ADODB.Recordset
Set ADObankRS = New ADODB.Recordset
ADObankRS.Open "SELECT [BANK ACCT ID],[BANK ACCT Balance] FROM [BANK Accounts] WHERE [BANK ACCT ID]='" & AccountNo & "'", db, adOpenKeyset, adLockOptimistic, adCmdText

If ADObankRS.RecordCount > 0 Then
        If DebitAmount <> 0 Then
          ADObankRS![BANK ACCT Balance] = ADObankRS![BANK ACCT Balance] + DebitAmount
        End If
        If CreditAmount <> 0 Then
          ADObankRS![BANK ACCT Balance] = ADObankRS![BANK ACCT Balance] - CreditAmount
        End If
        ADObankRS.Update
End If

ADObankRS.Close
Set ADObankRS = Nothing
End Sub

