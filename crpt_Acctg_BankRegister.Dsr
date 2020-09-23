VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} crpt_Acctg_BankRegister 
   ClientHeight    =   6315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11235
   OleObjectBlob   =   "crpt_Acctg_BankRegister.dsx":0000
End
Attribute VB_Name = "crpt_Acctg_BankRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Sub Section5_Format(ByVal pFormattingInfo As Object)
    Dim Conn As New ADODB.Connection
    Dim Rst As New ADODB.Recordset
    
    Conn.Open gblADOProvider
    Conn.CursorLocation = adUseClient
    
    Rst.Open "[rpt - Acctg - Bank Register]", Conn, adOpenForwardOnly, adLockReadOnly, _
        adCmdStoredProc
    
    Dim AccountDeposit As Currency, AccountWithdrawal As Currency
    
    While Not Rst.EOF
        If (Rst.Fields("Type") = "Cash Receipt" Or _
            Rst.Fields("Type") = "Deposit" Or _
            Rst.Fields("Type") = "Deposit Slip" Or _
            Rst.Fields("Type") = "Transfer To" And _
            Rst.Fields("BANK ID") = Field2.Value) Then
            AccountDeposit = AccountDeposit + Rst.Fields("Amount")
        End If
        Rst.MoveNext
    Wend
End Sub
