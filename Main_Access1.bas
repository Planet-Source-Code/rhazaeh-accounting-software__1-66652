Attribute VB_Name = "Main_Access"
Option Explicit

'Global Connection Strings
Global gblApplicationConnectString As String
Global gblADOProvider As String
Global gblBasicADOProvider As String
Global AppLoginName As String
Global CloseAllActive As Boolean
Global Mainstatus As Boolean
'Global dbBase As ADODB.Connection

Global gReportSelect$
Global intPeriod1 As Integer
Global intPeriod2 As Integer
Global intPeriod3 As Integer
Global intPeriod4 As Integer

Global gLinesPosted%
Global gNewInvoice$
Global gCustomerPeriod1Balance#
Global gCustomerPeriod2Balance#
Global gCustomerPeriod3Balance#
Global gCustomerPeriod4Balance#
Global gCustomerTotalBalance#

Global gNextCheckNo$               '<<<---- i used this

Global gREA$                       'Retained earnings account
Global gLastDayOfYear As Variant
Global gCustomerID$                'Used to pass ID from Customer Lookup Form
Global gCustomerdrill$  '----------not used 'check it'
Global gCOAdrill$       '----------not used 'check it'
Global gProjdrill       '----------not used 'check it'

Global gVendorID$

Global grdOnAddNew As Boolean

'--------------------------------- check this with option explicit
Global intAgeBy As Integer
Global gSalesID$
Global gReceiptID$
Global gPOID$
Global gPaymentID$
Global gPaymentAcct$
Global gInvID$
Global gCurrentReport$
Global gCurrentPrefix$
Global gWhere$
Global gPeriods As Integer
Global gPChoice As Integer
Global gMessage$
Global gHelpTopic$
Global gExiting%
Global gPayManyAmount@
Global gPayManyVendorID$
Global gPreview$
Global gAcctID$
Global gAcctLookType$
Global gShipToID$
Global gProductionID$
Global gProjectID$
Global gInventoryID$
Global gInvoiceID$
Global gPurchaseID$
Global gTaxID$
Global gTaxGroupID$
Global gBankDocument$
Global gEmployeeID$
Global gAdjustmentID$
Global gGLDocument$
Global gVendorPeriod1Balance#
Global gVendorPeriod2Balance#
Global gVendorPeriod3Balance#
Global gVendorPeriod4Balance#
Global gVendorTotalBalance#
Global gItemType$
Global gBankID$
Global gCustomerLookup$
Global gVendorLookup$
Global gItemLookup$
Global gAccountLookup$
Global gTaxLookup$
Global gTaxGroupLookup$
Global gShipToLookup$
Global gSalesLookup$
Global gOrderLookup$
Global gPurchaseLookup$
Global gARCheckLookup$
Global gAPCheckLookup$
Global gEmployeeLookup$
Global gGLLookup$
Global gInvAdjLookup$
Global gInvProLookup$
Global gProjectLookup$
Global gBankTransLookup$
Global gTCLookup$


Public fMainForm As Main_Menu

Public Sub DbConnectionString(UserInputdata As String)
Dim Passwd As String
    'Set all of the Global Parameters and Database Views
    'First Set Default Connection String & Provider
    'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\iabvb\project\database\default.mdb
    If UserInputdata = "" Or UserInputdata = App.Path & "\db1.mdb" Then
    Dim FileExist As String
        Passwd = "kucing"
        FileExist = Dir(App.Path & "\db1.mdb")
        If FileExist = "" Then
            MsgBox "Someone has deleted db1.mdb, the basic file." & vbCr & "Request the file from TBS. The application will stop executing", vbCritical, "Critical Error"
            End
        End If
        gblBasicADOProvider = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\db1.mdb"
        gblApplicationConnectString = UserInputdata
        gblADOProvider = gblBasicADOProvider
    Else
        gblApplicationConnectString = UserInputdata
        gblADOProvider = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & gblApplicationConnectString & ";Persist Security Info=False"
    End If
End Sub

Sub Main()
'Dim Passwd As String
'check if the application already running
If App.PrevInstance Then
      MsgBox App.EXEName & " already running!", 4096, "Warning"
      End
End If

'Dim FileExist As String
    
    Mainstatus = True
    'FileExist = Dir(App.Path & "\db1.mdb")
    'If FileExist = "" Then
    '    MsgBox "Someone has deleted db1.mdb, the basic file." & vbCr & "Request the file from TBS. The application will stop executing", vbCritical, "Critical Error"
    '    End
    'End If
    
    'Passwd = "kucing"
    'gblBasicADOProvider = App.Path & "\db1.mdb"
    DbConnectionString ""
    'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\NoSecret\iabvb\Project\Database\Default.mdb;Mode=Share Deny Read|Share Deny Write;Persist Security Info=True
    'gblBasicADOProvider = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
    'gblBasicADOProvider & ";Jet OLEDB:Database Password=" & Passwd & ";Admin"
    'gblBasicADOProvider = "Provider=Microsoft.Jet.OLEDB.4.0;Password=kucing;Data Source=" & gblBasicADOProvider & ";Persist Security Info=False"
    
    'gblBasicADOProvider = gblADOProvider
    'Load the Login Form
    'Start the Application, First load the Main
    Set fMainForm = New Main_Menu
    Load fMainForm
    
    Dim fLogin As New Main_Login
    fLogin.Show vbModal
    If Not fLogin.OK Then
        ExitApp
        End
    End If
    Unload fLogin
    
    fMainForm.Show

    CloseAllActive = False
End Sub

Sub SaveCompany()
ShowStatus True
Dim cnn As ADODB.Connection
Dim rsgetCompany As ADODB.Recordset
Dim SQLstatement As String
  'On Error GoTo GetCompany_Error

  Set cnn = New ADODB.Connection
  cnn.CursorLocation = adUseClient
  'cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\db1.mdb" & ";Persist Security Info=False"
  cnn.Open gblBasicADOProvider
  
  Set rsgetCompany = New ADODB.Recordset
  SQLstatement = "SELECT [Last Company],[date],[Last User] FROM [Last Company] ORDER BY [date]"
  With rsgetCompany
    .Open SQLstatement, cnn, adOpenKeyset, adLockOptimistic, adCmdText
    If .RecordCount < 20 Then
        .AddNew
        ![Last Company] = gblApplicationConnectString
        ![Date] = Format(Now, "mm/dd/yyyy hh:mm:ss")
        ![Last User] = AppLoginName
        .Update
    Else
        .MoveFirst
        ![Last Company] = gblApplicationConnectString
        ![Date] = Format(Now, "mm/dd/yyyy hh:mm:ss")
        ![Last User] = AppLoginName
        .Update
    End If
  End With
  
  'check the checklist
  IntroCheckList 1
  
  rsgetCompany.Close
  Set rsgetCompany = Nothing
  cnn.Close
  Set cnn = Nothing
  ShowStatus False
  Exit Sub
GetCompany_Error:
  Call ErrorLog("Menu Bar Module", "GetCompany", Now, Err.Number, Err.Description, True, cnn)
  Resume Next
  

End Sub

Sub LoadResStrings(frm As Form)
    'On Error Resume Next
End Sub

Sub GetCompany()
Dim cnn As ADODB.Connection
Dim rsgetCompany As ADODB.Recordset
Dim SQLstatement As String
Dim FileExist As Boolean
Dim FileName As String
Dim Response As Integer

On Error GoTo GetCompany_Error
ShowStatus True
  Set cnn = New ADODB.Connection
  cnn.CursorLocation = adUseServer
  cnn.Open gblBasicADOProvider
  
  Set rsgetCompany = New ADODB.Recordset
  SQLstatement = "SELECT [Last Company],[date],[Last User] FROM [Last Company] ORDER BY [date] ASC"
  With rsgetCompany
    .Open SQLstatement, cnn, adOpenStatic, adLockReadOnly, adCmdText
    If .RecordCount > 0 Then
        .MoveLast
        FileExist = False
        Do While Not FileExist
            FileName = Dir(![Last Company])
            'MsgBox ![Last Company]
            If FileName = "" Then
                FileExist = False
                'Response = MsgBox("The working table " & ![Last Company] & " Cannot be found" _
                '& vbCr & "Would you like computer to look for previous working Database or Exit", vbYesNo, "Loading The Previous Working Database")
                'If Response = vbNo Then End
            Else
                FileExist = True
                Mainstatus = True
                DbConnectionString ![Last Company]
            End If
            .MovePrevious
            If .BOF Then
                'MsgBox "There is no existing data or this may be a new setup" & vbCr & "Please use open or create new company for File", vbInformation, "Information"
                Mainstatus = False
                Exit Do
            End If
        Loop
  Else
    MsgBox "There is no existing data or this may be a new setup" & vbCr & "Please use open or create new company", vbInformation, "Information"
    Mainstatus = False
  End If
  End With
    
  rsgetCompany.Close
  Set rsgetCompany = Nothing
  cnn.Close
  Set cnn = Nothing
ShowStatus False
  Exit Sub
GetCompany_Error:
  'ErrorLog "Menu Bar Module", "GetCompany", Now, Err.Number, Err.Description, True, cnn
  End
  'Resume Next
  
End Sub

Function ExitApp() As Long
'Upload forms, revome all objects from memory
     Dim i As Integer
     Dim lngRetVal As Long
     '  Initialize routine
     'On Error GoTo ExitApp_EH1
     ExitApp = 0  '  assume success
     lngRetVal = 0
     
     ' Unload all forms
     For i = Forms.count - 1 To 0 Step -1
            Unload Forms(i)
     Next

ExitApp_Exit:
     'On Error Resume Next
     Close       '  release any file locks
     DoEvents
     ExitApp = lngRetVal
     Exit Function
ExitApp_EH1:
       lngRetVal = Err
       Resume ExitApp_Exit
End Function
 
 
