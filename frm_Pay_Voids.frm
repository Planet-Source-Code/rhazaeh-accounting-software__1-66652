VERSION 5.00
Begin VB.Form frm_Pay_Voids 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Void Checks"
   ClientHeight    =   3900
   ClientLeft      =   5160
   ClientTop       =   3810
   ClientWidth     =   4095
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   4095
   Begin VB.CommandButton cmdVoid 
      Caption         =   "Void Checks"
      Enabled         =   0   'False
      Height          =   855
      Left            =   2880
      Picture         =   "frm_Pay_Voids.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton cmdCheckNo 
      Height          =   255
      Left            =   3600
      Picture         =   "frm_Pay_Voids.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton cmdEmpID 
      Height          =   255
      Left            =   3600
      Picture         =   "frm_Pay_Voids.frx":0454
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton cmdBankAcct 
      Height          =   255
      Left            =   3600
      Picture         =   "frm_Pay_Voids.frx":059E
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox txtFields 
      DataField       =   " "
      DataSource      =   "adoPrimaryRS"
      Height          =   285
      Index           =   0
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   600
      Width           =   2055
   End
   Begin VB.TextBox txtFields 
      DataField       =   " "
      DataSource      =   "adoPrimaryRS"
      Height          =   285
      Index           =   1
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1200
      Width           =   2055
   End
   Begin VB.TextBox txtFields 
      DataField       =   " "
      DataSource      =   "adoPrimaryRS"
      Height          =   285
      Index           =   2
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1920
      Width           =   2055
   End
   Begin VB.TextBox txtFields 
      DataField       =   " "
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "M/d/yy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      DataSource      =   "adoPrimaryRS"
      Height          =   285
      Index           =   3
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   2280
      Width           =   2415
   End
   Begin VB.TextBox txtFields 
      DataField       =   " "
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      DataSource      =   "adoPrimaryRS"
      Height          =   285
      Index           =   4
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   2640
      Width           =   2415
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      Index           =   5
      Left            =   2760
      TabIndex        =   18
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "VOID"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   975
      Left            =   240
      TabIndex        =   17
      Top             =   2880
      Visible         =   0   'False
      Width           =   2505
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Void Checks"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   16
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label lblEmp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   14
      Top             =   1515
      Width           =   3615
   End
   Begin VB.Label lblAccts 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   900
      Width           =   3615
   End
   Begin VB.Label lblLabels 
      Caption         =   "Bank Account"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   9
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Amount"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label lblLabels 
      Caption         =   "Employee ID"
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   7
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Check No"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Check Date"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   1335
   End
End
Attribute VB_Name = "frm_Pay_Voids"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ADOprimaryrs As ADODB.Recordset
Attribute ADOprimaryrs.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean

Dim db As ADODB.Connection
'The recordset should only contain one record holding information pertaining to  a specific company
' inventory setup.


Private Sub cmdBankAcct_Click()
    Dim SQLstatement As String
    Dim No As Integer
    Dim ghead As String
    Dim fhead As String

    ClearText
    
    No = 1520
    SQLstatement = "select [BANK ACCT ID], [BANK ACCT Name]" & _
                    "from [BANK Accounts]"
    ghead = "Account Description"
    fhead = "Account No//Description"
    
    AllLookup.GetWhichTable No, SQLstatement, ghead, fhead, db
    'AllLookup.Show vbModal

End Sub

Private Sub ClearText()
    cmdVoid.Enabled = False
    Label2.Visible = False
    txtfields(2) = ""
    txtfields(3) = ""
    txtfields(4) = ""
    txtfields(5) = ""
End Sub

Private Sub cmdCheckNo_Click()
If txtfields(0) = "" Then
    MsgBox "Please select Bank Account before you could continue.", vbCritical, "Error"
    Exit Sub
End If
If txtfields(1) = "" Then
    MsgBox "Please select Employee ID before you could continue.", vbCritical, "Error"
    Exit Sub
End If

    Dim SQLstatement As String
    Dim No As Integer
    Dim ghead As String
    Dim fhead As String

    No = 1540
'    SQLStatement = "SELECT [AP PAY Check No], Format(Format([AP PAY Transaction Date],'mm/dd/yyyy'),'@@@@@@@@@@'), FORMAT(FORMAT([AP PAY Amount],'$###,###,##0.00'),'@@@@@@@@@@@@@'),[AP PAY Vendor No] FROM [AP Payment Header] WHERE [AP PAY Type] = 'Payroll'"
'    SQLStatement = SQLStatement & " AND [AP PAY Bank Account]='" & txtFields(0) & "' AND [AP PAY Vendor No]='" & txtFields(1) & "' ORDER BY [AP PAY Check No]"
    SQLstatement = "SELECT [AP PAY Check No],[AP PAY Transaction Date],[AP PAY Amount],[AP PAY Vendor No],[AP PAY Void] FROM [AP Payment Header] WHERE [AP PAY Type] = 'Payroll'"
    SQLstatement = SQLstatement & " AND [AP PAY Bank Account]='" & txtfields(0) & "' AND [AP PAY Vendor No]='" & txtfields(1) & "'"
    'Debug.Print SQLStatement
    ghead = "Employee"
    fhead = "Check No//Transaction Date//Amount//Employee Name//Status"
    
    AllLookup.GetWhichTable No, SQLstatement, ghead, fhead, db
    'AllLookup.Show vbModal
    If txtfields(5).Text = "-1" Then
        cmdVoid.Enabled = False
        Label2.Visible = True
    ElseIf txtfields(5).Text = "0" Then
        cmdVoid.Enabled = True
        Label2.Visible = False
    End If
End Sub


Private Sub cmdEmpID_Click()
    Dim SQLstatement As String
    Dim No As Integer
    Dim ghead As String
    Dim fhead As String
    
    ClearText

    No = 1530
    SQLstatement = "select [EMP ID], [EMP Name]" & _
                    "from [EMP Employees]"
    ghead = "Employee"
    fhead = "ID//Name"
    
    AllLookup.GetWhichTable No, SQLstatement, ghead, fhead, db
    'AllLookup.Show vbModal
End Sub

Private Sub cmdVoid_Click()
  'On Error GoTo cmdVoid_Click_Error

  Dim Response%
  
  If txtfields(2) = "" Then
  MsgBox "Enter a check to be voided", vbInformation, "Information"
  Exit Sub
  End If
  
  Response% = MsgBox("Are you sure you want to void this payment?", vbYesNo, "Confirmation")
  If Response% = vbNo Then Exit Sub
  
  'Proceed with void
  ShowStatus True
  Dim Success%

  db.BeginTrans
    
    Success% = VoidPayrollCheck(CStr(txtfields(1)), CStr(txtfields(2)), txtfields(3))
    
    If Success% = False Then
      db.RollbackTrans
      
      MsgBox "Transaction NOT Posted."
    Else
      db.CommitTrans
      
      MsgBox "Transaction Posted."
      'Me![VOID].Visible = True
      'DoCmd.GoToControl "AP PAY Check No"
    End If

  ShowStatus False
  
  Exit Sub

cmdVoid_Click_Error:
  Call ErrorLog("Pyrl - Voids", "cmdVoid_Click", Now, Err.Number, Err.Description, True, db)
  db.RollbackTrans
  ShowStatus False
  MsgBox ("Transaction Not Posted")
  Resume Next

End Sub

Private Sub Form_Load()
On Error GoTo FormErr
  Set db = New ADODB.Connection
  db.CursorLocation = adUseClient
  db.Open gblADOProvider

  Set ADOprimaryrs = New ADODB.Recordset
  'ADOprimaryRS.Open "select [SYS COM Inventory Adjustment Number],[SYS COM Inventory Cost Digits]," & _
    "[SYS COM Inventory Cost Method Last YN],[SYS COM Inventory Production Number]," & _
    "[SYS COM Inventory Qty Digits],[SYS COM Items per Transaction] from [SYS Company]", db, adOpenStatic, _
    adLockOptimistic

'  Dim oText As TextBox
  'Bind the text boxes to the data provider
'  For Each oText In Me.txtFields
'    Set oText.DataSource = ADOprimaryRS
'  Next
  
  GetTextColor Me
  mbDataChanged = False
Exit Sub
FormErr:
  MsgBox Err.Description
  ShowStatus False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If mbAddNewFlag Then Exit Sub

  Select Case KeyCode
    Case vbKeyEscape
      Unload Me
  End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo FormErr
  ShowStatus True

      If ADOprimaryrs.State = 0 Then GoTo JumpSkip
      If ADOprimaryrs.RecordCount > 0 Then
        If ADOprimaryrs.EditMode <> 0 Then
          ADOprimaryrs.CancelUpdate
        End If
      End If
      ADOprimaryrs.Close
      Set ADOprimaryrs = Nothing
JumpSkip:
      db.Close
      Set db = Nothing
  ShowStatus False
  Exit Sub
FormErr:
  MsgBox Err.Description
  ShowStatus False
End Sub

Private Sub adoPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This will display the current record position for this recordset
  'lblStatus.Caption = "Record: " & CStr(adoPrimaryRS.AbsolutePosition)
End Sub

Private Sub adoPrimaryRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This is where you put validation code
  'This event gets called when the following actions occur
  Dim bCancel As Boolean

  Select Case adReason
  Case adRsnAddNew
  Case adRsnClose
  Case adRsnDelete
  Case adRsnFirstChange
  Case adRsnMove
  Case adRsnRequery
  Case adRsnResynch
  Case adRsnUndoAddNew
  Case adRsnUndoDelete
  Case adRsnUndoUpdate
  Case adRsnUpdate
  End Select

  If bCancel Then adStatus = adStatusCancel
End Sub

Private Function VoidPayrollCheck(EmployeeID$, CheckNo$, CHECKDATE)

'On Error GoTo PostVoidError
  
  Dim CurrentBalance@
  Dim msg$
  Dim title$
  Dim rsPayment As ADODB.Recordset
  Dim rsCompany As ADODB.Recordset
  Dim rsEmployees As ADODB.Recordset
  Dim rsRegister As ADODB.Recordset
  Dim TranDate As Variant
  Dim TmpBalance@
  Dim Success%
  'Dim d As Database
  'Set d = CurrentDb
  
  Set rsPayment = New ADODB.Recordset
  rsPayment.Open "SELECT [AP PAY Check No],[AP PAY Transaction Date],[AP PAY Amount],[AP PAY Vendor No],[AP PAY Void] FROM [AP Payment Header] WHERE [AP PAY Vendor No]='" & EmployeeID$ & "' AND [AP PAY Check No]='" & CheckNo$ & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
  With rsPayment
  '.Index = "PrimaryKey"
  '.Seek "=", EmployeeID$, CheckNo$            'Seek record to void
  If .RecordCount = 0 Then
    MsgBox " This check number may be invalid.", vbCritical, "Error"
    VoidPayrollCheck = False
    'Debug.Print rsPayment.Source
    Exit Function
  Else
  '.Edit
  ![AP PAY Void] = -1
  .Update
  End If
  End With
  
  'Post by 1 - system date or 2 - Transaction date?
  Dim PostDate%
  Set rsCompany = New ADODB.Recordset
  rsCompany.Open "[SYS Company]", db, adOpenKeyset, adLockOptimistic, adCmdTable
  rsCompany.MoveFirst
  PostDate% = rsCompany("SYS COM GL Post By Date")

  'Set Post Date
  If PostDate% = 1 Then
    TranDate = DateValue(Format(Now, "Short Date"))
  Else
    TranDate = DateValue(rsPayment("AP PAY Transaction Date"))
  End If
  
  'Verify period can be posted to; Send TranDate; Return PeriodToPost and PeriodClosed
  Dim PeriodToPost%
  Dim PeriodClosed%
  Call VerifyPeriod(TranDate, PeriodToPost%, PeriodClosed%, db)
  If PeriodClosed% = True Then
    MsgBox "Unable to post transaction to a closed period.", , "Post Payment Error"
    VoidPayrollCheck = False
    Exit Function
  End If

  
  ' write GL Transaction Header
  Dim NewNumber&
  Dim SQLstatement As String
  Dim TempSQL As String
  
  SQLstatement = "INSERT INTO [GL Transaction]"
  SQLstatement = SQLstatement + " ([GL TRANS Document #],[GL TRANS Type],[GL TRANS Date],[GL TRANS Reference],"
  SQLstatement = SQLstatement + "[GL TRANS Amount],[GL TRANS Posted YN],[GL TRANS Description],"
  SQLstatement = SQLstatement + "[GL TRANS Source],[GL TRANS System Generated]) "
  TempSQL = "Payroll Void " & CheckNo$
  SQLstatement = SQLstatement + "VALUES ('" & TempSQL & "','Void Check', "
    If PostDate% = 1 Then   ' gl post date
      TempSQL = FormatDate(Now)
    Else
     TempSQL = rsPayment("AP PAY Transaction Date")
    End If
  SQLstatement = SQLstatement + TempSQL & ",'" & CheckNo$ & "'," & rsPayment("AP PAY Amount") & ",1,"
  TempSQL = "Payroll Entry " & CheckNo$
  SQLstatement = SQLstatement & "'" & TempSQL & "','Payroll',True)"
  'Debug.Print SQLstatement
  
  db.Execute SQLstatement
  
  Dim rsGLTrans As ADODB.Recordset
  Set rsGLTrans = New ADODB.Recordset
  rsGLTrans.Open "SELECT [GL TRANS Number] FROM [GL Transaction] WHERE [GL TRANS Document #]='" & "Payroll Void " & CheckNo$ & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
  NewNumber& = rsGLTrans![GL TRANS Number]
  rsGLTrans.Close
  Set rsGLTrans = Nothing
  
' rsGLTrans.AddNew
'    rsGLTrans("GL TRANS Document #") = "Payroll Void " & CheckNo$
'    rsGLTrans("GL TRANS Type") = "Void Check"
'    If PostDate% = 1 Then   ' gl post date
'      rsGLTrans("GL TRANS Date") = Format(Now, "Short Date") '------
'    Else
'      rsGLTrans("GL TRANS Date") = rsPayment("AP PAY Transaction Date")
'    End If
'    rsGLTrans("GL TRANS Reference") = CheckNo$
'    rsGLTrans("GL TRANS Amount") = rsPayment("AP PAY Amount")
'    rsGLTrans("GL TRANS Posted YN") = 1
'    rsGLTrans("GL TRANS Description") = "Payroll Entry " & CheckNo$
'    rsGLTrans("GL TRANS Source") = "Payroll"
'    rsGLTrans("GL TRANS System Generated") = True
'  rsGLTrans.Update
 
 'write to GL Detail Table
    Dim cmdVoided As Command
    Dim rsVoided As ADODB.Recordset
    Dim ParamVoided1 As Parameter
    Dim ParamVoided2 As Parameter
    
    Set cmdVoided = New Command
    cmdVoided.ActiveConnection = db
    
    cmdVoided.CommandText = "[Pyrl - SumGLDetailWorkVoided]"
    cmdVoided.CommandType = adCmdStoredProc
    
    Set ParamVoided1 = cmdVoided.CreateParameter("CheckNumber", adInteger, adParamInput) 'Screen.ActiveForm.[AP PAY Check No]       'set query criteria for current work table records
    Set ParamVoided2 = cmdVoided.CreateParameter("BankAcctNumber", adBSTR, adParamInput) 'Screen.ActiveForm.[AP PAY Check No]       'set query criteria for current work table records
    
    ParamVoided1.Value = txtfields(2).Text
    ParamVoided2.Value = txtfields(0).Text
    
    cmdVoided.Parameters.Append ParamVoided1
    cmdVoided.Parameters.Append ParamVoided2
    
    Set rsVoided = cmdVoided.Execute
    
    'VOIDED.Parameters![BankAcctNumber] = Screen.ActiveForm.[AP PAY Bank Account]
    'Set rsVoided = VOIDED.OpenRecordset
    
   Dim rsGLTransDetail As ADODB.Recordset
   Set rsGLTransDetail = New ADODB.Recordset
   rsGLTransDetail.Open "[GL Transaction Detail]", db, adOpenKeyset, adLockOptimistic, adCmdTable
    
    With rsGLTransDetail
    rsVoided.MoveFirst
    Do Until rsVoided.EOF
        .AddNew
        ![GL TRANSD Number] = NewNumber&
        ![GL TRANSD Account] = rsVoided!Account & ""
        ![GL TRANSD Debit Amount] = rsVoided!Credit 'Reverse
        ![GL TRANSD Credit Amount] = rsVoided!Debit 'Reverse
        ![GL TRANSD Project] = 0
        .Update
        
        'post GL Entry
         If rsVoided!Debit > 0 Then
            If rsVoided!Credit > 0 Then
                MsgBox ("Debit or Credit Data Error in Query 'SumGLDetailWorkVoided'")      'Error
                GoTo PostVoidError
            Else
                Success% = PostCOA(rsVoided!Account, TranDate, 0, rsVoided!Debit) 'Post Debit reversal
            End If
        End If
        
        If rsVoided!Credit > 0 Then
            If rsVoided!Debit > 0 Then
                MsgBox ("Debit or Credit Data Error in Query 'SumGLDetailWorkVoided'")      'Error
                GoTo PostVoidError
            Else
                Success% = PostCOA(rsVoided!Account, TranDate, rsVoided!Credit, 0) 'Post Credit reversal
            End If
        End If
        
        
        If Success% = False Then GoTo PostVoidError
        rsVoided.MoveNext
            If rsVoided.EOF Then
              Exit Do
            End If
    Loop
    End With
        
 ' Reduce Employee YTD Amounts


Dim sql As String
sql = "SELECT * FROM [Pyrl - Register] WHERE ((([Pyrl - Register].[EMP ID])= '" & EmployeeID$ & "') AND (([Pyrl - Register].CHECKDATE)= #" & Trim(CHECKDATE) & "#) AND (([Pyrl - Register].CHECKNUMBER)=" & Val(CheckNo$) & "))"
Set rsRegister = New ADODB.Recordset
rsRegister.Open sql, db, adOpenKeyset, adLockOptimistic, adCmdText
If rsRegister.RecordCount = 0 Then
       MsgBox ("Check not found in table 'Pyrl - Register'")
       VoidPayrollCheck = False
       Exit Function
End If

If rsRegister.RecordCount > 1 Then
       MsgBox ("Duplicate Checks in Table 'Pyrl - Register'")
       VoidPayrollCheck = False
       Exit Function
End If

Set rsEmployees = New ADODB.Recordset
rsEmployees.Open "Select * FROM [Pyrl - Employee Data] WHERE [EMP ID] = '" & EmployeeID$ & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
If rsEmployees.RecordCount = 0 Then
    MsgBox "Employee " & EmployeeID$ & " could not be found in table 'Pyrl - Employee Data'." & Chr$(10) & "Year-to-Date Gross Pay could not be reduced."
    VoidPayrollCheck = False
    Exit Function
Else
  With rsEmployees
  '.Edit
  ![YTDGROSS] = ![YTDGROSS] - rsRegister!GROSS
  ![YTDFICA] = ![YTDFICA] - rsRegister!FICA
  ![YTDFIT] = ![YTDFIT] - rsRegister!FIT
  ![YTDSTATETAX] = ![YTDSTATETAX] - rsRegister!STATETAX
  ![YTDLOCAL] = ![YTDLOCAL] - rsRegister!LOCAL
  ![YTDREGHOURS] = ![YTDREGHOURS] - rsRegister!REGHOURS
  ![YTDOTHOURS] = ![YTDOTHOURS] - rsRegister!OTHOURS
  .Update
  End With
End If
  
  'rsRegister.Edit
  rsRegister!VOIDED = -1   'void in Register
  rsRegister.Update

VoidPayrollCheck = True
Exit Function

PostVoidError:
  VoidPayrollCheck = False
  Call ErrorLog("Purchase Module", "PostVoid", Now, Err.Number, Err.Description, True, db)
  Resume Next
  Exit Function
 
End Function
