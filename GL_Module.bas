Attribute VB_Name = "GL_Module"

Function CloneGLEntry(DocumentKey&, db As ADODB.Connection) As Integer

  'On Error GoTo CloneGLEntry_Error
  
  'Dim db As ADODB.Connection
  'Set db = New ADODB.Connection
  'db.CursorLocation = adUseClient
  'db.Open gblADOProvider
  
  Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
  
  Dim rs2 As ADODB.Recordset
  Set rs2 = New ADODB.Recordset
  
  Dim rsDetail As ADODB.Recordset
  Set rsDetail = New ADODB.Recordset
  
  Dim rsDetail2 As ADODB.Recordset
  Set rsDetail2 = New ADODB.Recordset
  
  'Dim rsRecur As ADODB.Recordset
  'rsRecur.Open "[SYS Recurred]", db, adOpenStatic, adLockOptimistic, adCmdTable

  rs.Open "SELECT * FROM [GL Transaction] WHERE [GL TRANS Number]=" & DocumentKey&, db, adOpenKeyset, adLockOptimistic, adCmdTable
  rs2.Open "SELECT * FROM [GL Transaction] WHERE [GL TRANS Number]=" & DocumentKey&, db, adOpenKeyset, adLockOptimistic, adCmdTable
  'rs2.Open "[GL Transaction]", db, adOpenStatic, adLockOptimistic, adCmdTable
  
  'rs.Index = "PrimaryKey"
  'rs.Seek "=", DocumentKey&

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
    'MyCounter2& = rs2("GL Trans Number")
    'For X% = 0 To count% - 1
    '  rs2(X%) = rs(X%)
    'Next X%
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

    'rs2("GL Trans Number") = MyCounter2&
    'Rename Ext Document #
    'If AskForID% = True Then
    '  gNewInvoice$ = InputBox("Enter new document #")
    'Else
      'Create a document ID
    '  Dim rsSeek As ADODB.Recordset
    '  rsSeek.Open "GL Transaction", db, adOpenStatic, adLockOptimistic, adCmdTable
    '  rsSeek.Index = "GL TRANS Document #"
    '  Dim Counter%
    '  Counter% = 1
    '  Dim Success%
    '  Success% = False
    '  Do While Not Success%
    '    gNewInvoice$ = rs2("GL TRANS Document #") & "-" & Trim(Str(Counter%))
    '    'Check if this newly created document exists
    '    rsSeek.Seek gNewInvoice$
    '    If rsSeek.EOF Then
    '      Success% = True
    '    Else
    '      Success% = False
    '      Counter% = Counter% + 1
    '    End If
    '  Loop
    'End If
    'If gNewInvoice$ = "" Then
    '  db.RollbackTrans
    '  CloneGLEntry% = 1
    '  Exit Function
    'End If
    
    rs2("GL TRANS Date") = CDate(FormatDate(Date))
    rs2("GL TRANS Document #") = "ConvertGL" & AppLoginName
    rs2("GL TRANS Recurring YN") = False
    rs2("GL TRANS Posted YN") = False
    'rsRecur.AddNew
    '  rsRecur("Document Type") = "GL Entry"
    '  rsRecur("Document Number") = rs2("GL TRANS Document #")
    '  rsRecur("Reference") = rs2("GL TRANS Reference")
    '  rsRecur("Amount") = rs2("GL TRANS Amount")
    'rsRecur.Update

  rs2.Update
  rs2.Close
  Set rs2 = Nothing
  
  Set rs2 = New ADODB.Recordset
  rs2.Open "SELECT [GL TRANS Number],[GL TRANS Document #] FROM [GL Transaction] where [GL TRANS Document #] ='ConvertGL" & AppLoginName & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
    MyCounter2& = rs2![GL TRANS Number]
    rs2![GL TRANS Document #] = "[" & AppLoginName & Format(Now, "MMdd") & Right(Format(MyCounter2&, "0000"), 4) & "]"
  rs2.Update
  rs2.Close
  Set rs2 = Nothing

  Dim DetailCounter&

  rsDetail.Open "SELECT * FROM [GL Transaction Detail] where [GL TRANSD Number] = " & MyCounter&, db, adOpenKeyset, adLockOptimistic, adCmdText
  'On Error Resume Next
  'Err = 0
  'rsDetail.MoveLast
  If rsDetail.RecordCount = 0 Then
    'No Detail
  Else
    rsDetail.MoveFirst
    'Create new detail record
    'rsDetail2.Open "[GL Transaction Detail]", db, adOpenKeyset, adLockOptimistic, adCmdTable
    rsDetail2.Open "SELECT * FROM [GL Transaction Detail] where [GL TRANSD Number] = " & MyCounter&, db, adOpenKeyset, adLockOptimistic, adCmdText
    Do While Not rsDetail.EOF
      count% = rsDetail.Fields.count
      rsDetail2.AddNew
        'DetailCounter& = rsDetail2("GL TRANSD ID")
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
        'rsDetail2("GL TRANSD ID") = DetailCounter&
        rsDetail2("GL TRANSD Number") = MyCounter2&
        'Add rest of detail records
      rsDetail2.Update
      rsDetail.MoveNext
    Loop
  End If

SkipDetail:
  
  db.CommitTrans
  CloneGLEntry% = True
  rs2.Close
  Set rs2 = Nothing
  'rsRecur.Close
  'Set rsRecur = Nothing
  rsDetail.Close
  Set rsDetail = Nothing
  rsDetail2.Close
  Set rsDetail2 = Nothing
  rs.Close
  Set rs = Nothing
  'rsSeek.Close
  'Set rsSeek = Nothing
  'db.Close
  'Set db = Nothing
  Exit Function

CopyFailed:
  db.RollbackTrans
  CloneGLEntry% = False
  rs2.Close
  Set rs2 = Nothing
  'rsRecur.Close
  'Set rsRecur = Nothing
  rsDetail.Close
  Set rsDetail = Nothing
  rsDetail2.Close
  Set rsDetail2 = Nothing
  rs.Close
  Set rs = Nothing
  'rsSeek.Close
  'Set rsSeek = Nothing
  'db.Close
  'Set db = Nothing
  Exit Function

CloneGLEntry_Error:
  Call ErrorLog("GL Module", "CloneGLEntry", Now, Err.Number, Err.Description, True, db)
  Resume Next

  rs2.Close
  Set rs2 = Nothing
  'rsRecur.Close
  'Set rsRecur = Nothing
  rsDetail.Close
  Set rsDetail = Nothing
  rsDetail2.Close
  Set rsDetail2 = Nothing
  rs.Close
  Set rs = Nothing
  'rsSeek.Close
  'Set rsSeek = Nothing
  'db.Close
  'Set db = Nothing

End Function

Function RecurGL() As Integer

'  Dim db As ADODB.Connection
'  Set db = New ADODB.Connection
'  db.CursorLocation = adUseClient
'  db.Open gblADOProvider
  
'  Dim rsRecur As ADODB.Recordset
'  rsRecur.Open "SELECT * FROM [GL Transaction] where [GL TRANS Recurring YN] = true", db, adOpenStatic, adLockOptimistic, adCmdText

'  Dim DocumentKey&
'  Dim Success%
  'On Error Resume Next
'  rsRecur.MoveFirst
'  If rsRecur.EOF Then GoTo SkipRecurGL

'  db.BeginTrans

'  Do While Not rsRecur.EOF
'    DocumentKey& = rsRecur("GL TRANS Number")
'    Success% = CloneGLEntry(DocumentKey&, False)
'    rsRecur.MoveNext
'  Loop

'  db.CommitTrans

'SkipRecurGL:
'  rsRecur.Close
'  Set rsRecur = Nothing
'  db.Close
'  Set db = Nothing
'  Exit Function

'  rsRecur.Close
'  Set rsRecur = Nothing
'  db.Close
'  Set db = Nothing

End Function

Function ReverseGLEntry(DocumentKey&, AskForID%) As Integer

  'On Error GoTo ReverseGLEntry_Error

  Dim db As ADODB.Connection
  Set db = New ADODB.Connection
  db.CursorLocation = adUseServer
  db.Open gblADOProvider
   
  Dim rs As ADODB.Recordset
  Dim rs2 As ADODB.Recordset

  Dim rsDetail As ADODB.Recordset
  Dim rsDetail2 As ADODB.Recordset
  
  Set rs = New ADODB.Recordset
  rs.Open "[GL Transaction]", db, adOpenStatic, adLockOptimistic, adCmdTable
  Set rs2 = New ADODB.Recordset
  rs2.Open "[GL Transaction]", db, adOpenStatic, adLockOptimistic, adCmdTable

  'rs.Index = "PrimaryKey"
  'rs.Seek DocumentKey&
  rs.MoveFirst
  rs.Find "[GL TRANS Number]=" & DocumentKey&

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
    'MyCounter2& = rs2("GL Trans Number")
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

    'rs2("GL Trans Number") = MyCounter2&
    'Rename Ext Document #
    If AskForID% = True Then
    '  gNewInvoice$ = InputBox("Enter new document #")
    'Else
      'Create a document ID
      Dim rsSeek As ADODB.Recordset
      Set rsSeek = New ADODB.Recordset
      rsSeek.Open "[GL Transaction]", db, adOpenStatic, adLockOptimistic, adCmdTable
      'rsSeek.Index = "GL TRANS Document #"
      Dim Counter%
      Counter% = 1
      Dim Success%
      Success% = False
      Do While Not Success%
        gNewInvoice$ = rs2("GL TRANS Document #") & "-" & Trim(Str(Counter%))
        'Check if this newly created document exists
        rsSeek.MoveFirst
        rsSeek.Find "[GL TRANS Document #]='" & gNewInvoice$ & "'"
        If rsSeek.EOF Then
          Success% = True
          MsgBox "Your reverse transaction number is " & gNewInvoice$, vbInformation, "Information"
        Else
          Success% = False
          Counter% = Counter% + 1
        End If
      Loop
    End If
    If gNewInvoice$ = "" Then
      db.RollbackTrans
      ReverseGLEntry% = 1
      Exit Function
    End If
    rs2("GL TRANS Document #") = gNewInvoice$
    rs2("GL TRANS Recurring YN") = False
    rs2("GL TRANS Posted YN") = False
  rs2.Update
  
  MyCounter2& = rs2("GL Trans Number")

  Dim DetailCounter&
  Set rsDetail = New ADODB.Recordset
  rsDetail.Open "SELECT * FROM [GL Transaction Detail] where [GL TRANSD Number] = " & MyCounter&, db, adOpenStatic, adLockOptimistic, adCmdText
  'On Error Resume Next
  Err = 0
  'rsDetail.MoveLast
  'rsDetail.MoveFirst
  If rsDetail.RecordCount = 0 Then
    'No Detail
  Else
    'Create new detail record
    Set rsDetail2 = New ADODB.Recordset
    rsDetail2.Open "[GL Transaction Detail]", db, adOpenStatic, adLockOptimistic, adCmdTable
    Do While Not rsDetail.EOF
      count% = rsDetail.Fields.count
      'rsDetail2.CancelUpdate
      rsDetail2.AddNew
        'DetailCounter& = rsDetail2("GL TRANSD ID")
        'rsDetail2("GL TRANSD ID") = DetailCounter&
        'rsDetail2("GL TRANSD Number") = MyCounter2&
        For X% = 0 To count% - 2
            If IsNull(rsDetail(X%)) = False Then
                If rsDetail(X%).Name = "GL TRANSD Debit Amount" Then
                  rsDetail2("GL TRANSD Credit Amount") = rsDetail("GL TRANSD Debit Amount")
                ElseIf rsDetail(X%).Name = "GL TRANSD Credit Amount" Then
                  rsDetail2("GL TRANSD Debit Amount") = rsDetail("GL TRANSD Credit Amount")
                Else
                  rsDetail2(X%) = rsDetail(X%)
                End If
            End If
        Next X%
        rsDetail2("GL TRANSD Number") = MyCounter2&
        'Add rest of detail records
      rsDetail2.Update
      rsDetail.MoveNext
    Loop
  End If

SkipDetail2:
  
  db.CommitTrans
  ReverseGLEntry% = True
  rs.Close
  Set rs = Nothing
  rs2.Close
  Set rs2 = Nothing
  rsDetail.Close
  Set rsDetail = Nothing
  rsDetail2.Close
  Set rsDetail2 = Nothing
  'rsSeek.Close
  Set rsSeek = Nothing
  db.Close
  Set db = Nothing
  Exit Function

ReverseFailed:
  db.RollbackTrans
  ReverseGLEntry% = False
  rs.Close
  Set rs = Nothing
  rs2.Close
  Set rs2 = Nothing
  rsDetail.Close
  Set rsDetail = Nothing
  rsDetail2.Close
  Set rsDetail2 = Nothing
  rsSeek.Close
  Set rsSeek = Nothing
  db.Close
  Set db = Nothing
  
  Exit Function
  
ReverseGLEntry_Error:
  Call ErrorLog("GL Module", "ReverseGLEntry", Now, Err.Number, Err.Description, True, db)
  Resume Next

  rs.Close
  Set rs = Nothing
  rs2.Close
  Set rs2 = Nothing
  rsDetail.Close
  Set rsDetail = Nothing
  rsDetail2.Close
  Set rsDetail2 = Nothing
  rsSeek.Close
  Set rsSeek = Nothing
  db.Close
  Set db = Nothing

End Function

