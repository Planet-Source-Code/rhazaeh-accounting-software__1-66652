Attribute VB_Name = "Utility_Module"


Function DropAllBut2(Amount@) As Currency
'what is the used of this function --- can it be eliminate???
  'On Error Resume Next
  Dim Longer&
  Dim Currencie@
  Currencie@ = Amount@ * 100
  Longer& = CLng(Currencie@)
  DropAllBut2 = Longer& / 100

End Function

Function GetUOMMultiplier(Item As Variant, Unit As Variant, db As ADODB.Connection) As Double

  Dim msg$
  Dim title$
  Dim Multiplier As Double

  'On Error GoTo GetUOMMultiplier_Error

  Multiplier = 0
  ' handle no item id
  If IsNull(Item) Then
    GoTo GetUOMMultiplier_Exit
  End If
    
  ' handle no item id
  If Len(Trim$(Item)) = 0 Then
    GoTo GetUOMMultiplier_Exit
  End If

  ' handle no units
  If IsNull(Unit) Then
    GoTo GetUOMMultiplier_Exit
  End If

  ' handle no units
  If Len(Trim$(Unit)) = 0 Then
    GoTo GetUOMMultiplier_Exit
  End If

  'Dim db As ADODB.Connection
  'Set db = New ADODB.Connection
  'db.CursorLocation = adUseClient
  'db.Open gblADOProvider
  Dim rsItemsBreak As ADODB.Recordset
  Set rsItemsBreak = New ADODB.Recordset
  rsItemsBreak.Open "SELECT * FROM [INV Items Break] WHERE [INV BREAK ID]='" & Item & "'", db, adOpenKeyset, adLockReadOnly, adCmdText
  
  ' go get the unit multiplier
  '------------->>>this is wrong, in the database the unit input were written in the [INV BREAK Description] instead
  'of [INV BREAK Unit]--- i have move the data to the proper place but i don't have time to find
  'the data input at this time, i hope if someone did see this message could inform me at razi@aumgrp.po.my
  'rsItemsBreak.Index = "IDDescription"
  'rsItemsBreak.Seek Item, Unit
  'rsItemsBreak.MoveFirst
  'rsItemsBreak.Find "[INV BREAK ID]='" & Item & "'" ' AND [INV BREAK Unit]='" & Unit & "'"
  If rsItemsBreak.RecordCount = 0 Then
    Multiplier = 1
  Else
    Multiplier = IIf(IsNull(rsItemsBreak("INV Break Qty")), 1, rsItemsBreak("INV Break Qty"))
  End If

GetUOMMultiplier_Exit:

  If Multiplier = 0 Then
    Multiplier = 1
  End If
  
  GetUOMMultiplier = Multiplier
  
 rsItemsBreak.Close
 Set rsItemsBreak = Nothing
 'db.Close
 'Set db = Nothing
   
  Exit Function

GetUOMMultiplier_Error:
  Call ErrorLog("Utility Module", "GetUOMMultiplier", Now, Err.Number, Err.Description, False, db)
  'Resume Next
  Resume GetUOMMultiplier_Exit

End Function

Function isloaded(FormName As String) As Boolean

  On Error GoTo IsLoaded_Error
    
  isloaded = False
  Dim checkform As Form
   
  For Each checkform In Forms
    If checkform.Name = FormName Then
        isloaded = True
        Exit For
    End If
  Next
      
  Exit Function
IsLoaded_Error:
  'Call ErrorLog("Utility Module", "IsLoaded", Now, Err.Number, Err.Description, False, db)
  Resume Next

End Function
Function isOpen(FormName As String, ObjectType As Integer)

  'On Error GoTo IsOpen_Error
    
  isOpen = False
  Dim checkform As Form
   
  For Each checkform In Forms
    If checkform.Name = FormName Then
        isOpen = True
        Exit For
    End If
  Next
      
  Exit Function
IsOpen_Error:
  'Call ErrorLog("Utility Module", "IsOpen", Now, Err.Number, Err.Description, False, db)
  Resume Next

End Function


Sub LockDetail(frm As Form)

  'On Error GoTo LockDetail_Error

  'Lock all controls in detail section
  
  Exit Sub
LockDetail_Error:
  'Call ErrorLog("Utility Module", "LockDetail", Now, Err.Number, Err.Description, True, db)
  Resume Next

End Sub

Sub ErrorLog(FormName$, ModuleName$, When As Variant, ErrorCode%, ErrorString$, DispMessage%, db As ADODB.Connection)

On Error GoTo ErrorLog_Error

  ShowStatus False

  If DispMessage% = True Then
    MsgBox "The following unexpected error has occurred:" & Chr$(10) & ErrorString$ & Chr$(10) & Chr$(10) & "Please check data and setup parameters.", , "Error"
  End If
    
  'Dim rs As ADODB.Recordset
  'Set rs = New ADODB.Recordset
  'rs.Open "[Error Log]", db, adOpenStatic, adLockOptimistic, adCmdTable
  
      SQLstatement = "INSERT INTO [Error Log]"
      SQLstatement = SQLstatement & " ([Form Name],[Module Name],[When],[Error Code],[Error String])"
      SQLstatement = SQLstatement & " VALUES ('" & FormName$ & "','" & ModuleName$ & "',#" & When & "#," & ErrorCode% & ",'" & Left$(ErrorString$, 255) & "')"
      db.Execute SQLstatement
  
  'rs.AddNew
  '  rs("Form Name") = FormName$
  '  rs("Module Name") = ModuleName$
  '  rs("When") = When
  '  rs("Error Code") = ErrorCode%
  '  rs("Error String") = Left$(ErrorString$, 255)
  'rs.Update
    
  ShowStatus False
  Exit Sub
  
ErrorLog_Error:
  MsgBox "You have lost your core system... Please contact TBS", , "Error"
  ShowStatus False
  Exit Sub

 
  Exit Sub
  
End Sub

Function Round#(Num#)

  'On Error GoTo Round_Error

  Dim X#
  Dim Y#
  Dim z#

  z# = Num# * 100
  
  X# = CLng(z#)
  Y# = z# - X#
  If Y# >= 0.5 Then
    X# = X# + 1
  End If

  X# = X# / 100

  Round# = X#

  Exit Function
Round_Error:
  'Call ErrorLog("Utility Module", "Round", Now, Err.Number, Err.Description, True, db)
  Resume Next

End Function

Sub UnlockDetail(frm As Form)
  
  'On Error GoTo UnlockDetail_Error
  
  'Lock all controls in detail section
  
  Exit Sub
UnlockDetail_Error:
  'Call ErrorLog("Utility Module", "UnlockDetail", Now, Err.Number, Err.Description, True, db)
  Resume Next

End Sub

Function ValidateKeyAscii(KeyAscii As Integer)
  
  'On Error GoTo ValidateKeyAscii_Error
  
  KeyAscii% = Asc(UCase(Chr(KeyAscii)))

  If KeyAscii = 32 Then KeyAscii = 0

  ValidateKeyAscii = KeyAscii
  
  Exit Function
ValidateKeyAscii_Error:
  'Call ErrorLog("Utility Module", "ValidateKeyAscii", Now, Err.Number, Err.Description, True, db)
  Resume Next

End Function


Function NumToString(valStr As String)
   
    'On Error GoTo NumToString_Error
   
    Static ones$(0 To 9)
    Static teens$(0 To 9)
    Static tens$(0 To 9)
    Static thousands$(0 To 4)
    Dim buff$
    Dim tmpBuff$
    Dim tmpStr$
    Dim tmpVar#
    Dim col%
    Dim aDigit%
    Dim i%
    Dim allZeros%

    ones$(0) = "zero"
    ones$(1) = "one"
    ones$(2) = "two"
    ones$(3) = "three"
    ones$(4) = "four"
    ones$(5) = "five"
    ones$(6) = "six"
    ones$(7) = "seven"
    ones$(8) = "eight"
    ones$(9) = "nine"

    teens$(0) = "ten"
    teens$(1) = "eleven"
    teens$(2) = "twelve"
    teens$(3) = "thirteen"
    teens$(4) = "fourteen"
    teens$(5) = "fifteen"
    teens$(6) = "sixteen"
    teens$(7) = "seventeen"
    teens$(8) = "eighteen"
    teens$(9) = "nineteen"

    tens$(0) = ""
    tens$(1) = "ten"
    tens$(2) = "twenty"
    tens$(3) = "thirty"
    tens$(4) = "forty"
    tens$(5) = "fifty"
    tens$(6) = "sixty"
    tens$(7) = "seventy"
    tens$(8) = "eighty"
    tens$(9) = "ninety"

    thousands$(0) = ""
    thousands$(1) = "thousand"
    thousands$(2) = "million"
    thousands$(3) = "billion"
    thousands$(4) = "trillion"
              
    'Trap errors
    'On Error GoTo NumToString_Error
    
    'Normalize number (Cdbl() will strip commas, leading zeros, etc. for us)
    tmpVar# = CDbl(valStr)
    
    'Get fractional part
    buff$ = "and " & Format((tmpVar# - Int(tmpVar#)) * 100, "00") & "/100"
    
    'Convert rest to string and process each digit
    tmpStr$ = CStr(Int(tmpVar#))
    
    'Iterate through string
    For i% = Len(tmpStr$) To 1 Step -1
        
        'Get value of this digit
        aDigit% = Val(Mid$(tmpStr$, i, 1))
        
        'Get column position
        col% = (Len(tmpStr$) - i) + 1
        
        'Action depends on 1's, 10's or 100's column
        Select Case (col% Mod 3)
            Case 1  '1's position
                allZeros% = False
                If i = 1 Then
                    tmpBuff$ = ones$(aDigit%) & " "
                ElseIf Mid$(tmpStr$, i - 1, 1) = "1" Then
                    tmpBuff$ = teens$(aDigit%) & " "
                    i = i - 1   'Skip tens position
                ElseIf aDigit% > 0 Then
                    tmpBuff$ = ones$(aDigit%) & " "
                Else
                    'If next 10s & 100s cols are also 0, don't show 'thousands'
                    allZeros% = True
                    If i > 1 Then
                        If Mid$(tmpStr$, i - 1, 1) <> "0" Then
                            allZeros% = False
                        End If
                    End If
                    If i > 2 Then
                        If Mid$(tmpStr$, i - 2, 1) <> "0" Then
                            allZeros% = False
                        End If
                    End If
                    tmpBuff$ = ""
                End If
                If allZeros% = False And col% > 1 Then
                    tmpBuff$ = tmpBuff$ & thousands$(col% / 3) & " "
                End If
                buff$ = tmpBuff$ & buff$
            Case 2  '10's position
                If aDigit% > 0 Then
                    buff$ = tens$(aDigit%) & " " & buff$
                End If
            Case 0  '100's position
                If aDigit% > 0 Then
                    buff$ = ones$(aDigit%) & " hundred " & buff$
                End If
        End Select
    Next i%
    
    'Convert first letter to upper case
    If Len(buff$) > 0 Then buff$ = UCase$(Left$(buff$, 1)) & Mid$(buff$, 2)

EndNumToString:
    
    'Return result
    NumToString = buff$
    
    Exit Function

NumToString_Error:
    'Call ErrorLog("Utilities", "NumToString", Now, Err.Number, Err.Description, True, db)
    Resume EndNumToString

End Function

Function ZeroNulltoBlank(AnyData As Variant) As Variant
                                    
  'On Error GoTo ZeroNullToBlank_Error
                                    
  ' this routine converts passed data to blanks if it is null or
  '   blank otherwise it returns a formatted dollar amount

  If IsNull(AnyData) Then
    ZeroNulltoBlank = ""
  ElseIf AnyData = 0 Then
    ZeroNulltoBlank = ""
  Else
    ZeroNulltoBlank = Format(AnyData, "Currency")
  End If
  
  Exit Function
ZeroNullToBlank_Error:
  'Call ErrorLog("Utilities", "ZeroNullToBlank", Now, Err.Number, Err.Description, True, db)
  Resume Next
    
End Function

Function GetLimit(InputData$) As Double

  'On Error GoTo GetLimit_Error

  InputData$ = StripQuotes(InputData$)
  
  GetLimit = Val(InputData$)
  
  Exit Function
GetLimit_Error:
  'Call ErrorLog("File Handling", "GetLimit", Now, Err.Number, Err.Description, True, db)
  Resume Next
  
End Function


Function StripQuotes(InputData$) As String

  'On Error GoTo StipQuotes_Error

  Dim GoodText$
  Dim X%
  
  X% = InStr(InputData$, Chr$(34))
  If X% > 0 Then
    InputData$ = Mid$(InputData$, X% + 1)
    X% = InStr(InputData$, Chr$(34))
    If X% > 0 Then
      InputData$ = Left$(InputData$, X% - 1)
    End If
  End If
  
  StripQuotes = InputData$
  
  Exit Function
StipQuotes_Error:
  'Call ErrorLog("File Handling", "StripQuotes", Now, Err.Number, Err.Description, True, db)
  Resume Next
  
End Function


Function TabLoc(InputData$, TabStop%) As String

  'Return string between desired tab stop and previous tab stop
  'Tab is chr(9)
  Dim InCharacter$
  Dim GoodText$
  Dim X%
  Dim TabCount%
  
  'On Error Resume Next
  
  GoodText$ = ""
  If TabStop% = 1 Then
    'Read up to first tab
    X% = 1
    InCharacter$ = Mid$(InputData$, X%, 1)
    Do While InCharacter$ <> Chr(9)
      GoodText$ = GoodText$ & InCharacter$
      X% = X% + 1
      Err = 0
      InCharacter$ = Mid$(InputData$, X%, 1)
      If X% > Len(InputData$) Then
        'Must be at end of string
        Exit Do
      End If
    Loop
  Else
    'Find location of tabstop -1
    TabCount% = 0
    X% = 1
    Do While TabCount% < TabStop% - 1
      InCharacter$ = Mid$(InputData$, X%, 1)
      Do While InCharacter$ <> Chr(9)
        X% = X% + 1
        Err = 0
        InCharacter$ = Mid$(InputData$, X%, 1)
        If X% > Len(InputData$) Then
          'Must be at end of string
          Exit Do
        End If
      Loop
      X% = X% + 1
      TabCount% = TabCount% + 1
    Loop
    'x% is now at TabCount% -1
    GoodText$ = ""
    InCharacter$ = Mid$(InputData$, X%, 1)
    Do While InCharacter$ <> Chr(9)
      GoodText$ = GoodText$ & InCharacter$
      X% = X% + 1
      Err = 0
      InCharacter$ = Mid$(InputData$, X%, 1)
      If X% > Len(InputData$) Then
        'Must be at end of string
        Exit Do
      End If
    Loop
  End If
  
  TabLoc$ = GoodText$
  
End Function
