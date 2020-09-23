VERSION 5.00
Begin VB.Form frm_AR_Cust_Financials 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Customer Financials"
   ClientHeight    =   4470
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   10485
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   10485
   Begin VB.TextBox txtFields 
      DataField       =   "AR CUST Customer ID"
      DataSource      =   "adoPrimaryRS"
      Height          =   285
      Index           =   0
      Left            =   1440
      TabIndex        =   24
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtFields 
      DataField       =   "AR CUST Name"
      DataSource      =   "adoPrimaryRS"
      Height          =   285
      Index           =   1
      Left            =   1440
      TabIndex        =   23
      Top             =   480
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      DataField       =   "AR CUST Financial Period 1"
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
      Index           =   2
      Left            =   840
      TabIndex        =   22
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF00&
      DataField       =   "AR CUST Financial Period 2"
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
      Index           =   3
      Left            =   2760
      TabIndex        =   21
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000FFFF&
      DataField       =   "AR CUST Financial Period 3"
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
      Left            =   4680
      TabIndex        =   20
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000000FF&
      DataField       =   "AR CUST Financial Period 4"
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
      Index           =   5
      Left            =   6600
      TabIndex        =   19
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      DataField       =   "AR CUST Financial Total"
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
      Index           =   6
      Left            =   8520
      TabIndex        =   18
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      DataField       =   "AR CUST Highest Balance"
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
      Index           =   7
      Left            =   8520
      TabIndex        =   17
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      DataField       =   "AR CUST Invoices Last Year"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      DataSource      =   "adoPrimaryRS"
      Height          =   285
      Index           =   8
      Left            =   4680
      TabIndex        =   16
      Top             =   2880
      Width           =   1815
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      DataField       =   "AR CUST Invoices Lifetime"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      DataSource      =   "adoPrimaryRS"
      Height          =   285
      Index           =   9
      Left            =   6600
      TabIndex        =   15
      Top             =   2880
      Width           =   1815
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      DataField       =   "AR CUST Invoices YTD"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      DataSource      =   "adoPrimaryRS"
      Height          =   285
      Index           =   10
      Left            =   2760
      TabIndex        =   14
      Top             =   2880
      Width           =   1815
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      DataField       =   "AR CUST Payment Terms"
      DataSource      =   "adoPrimaryRS"
      Height          =   285
      Index           =   11
      Left            =   1440
      TabIndex        =   13
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      DataField       =   "AR CUST Payments Last Year"
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
      Index           =   12
      Left            =   4680
      TabIndex        =   12
      Top             =   3240
      Width           =   1815
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      DataField       =   "AR CUST Payments Lifetime"
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
      Index           =   13
      Left            =   6600
      TabIndex        =   11
      Top             =   3240
      Width           =   1815
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      DataField       =   "AR CUST Payments YTD"
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
      Index           =   14
      Left            =   2760
      TabIndex        =   10
      Top             =   3240
      Width           =   1815
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      DataField       =   "AR CUST Sales Last Year"
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
      Index           =   16
      Left            =   4680
      TabIndex        =   9
      Top             =   2520
      Width           =   1815
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      DataField       =   "AR CUST Sales Lifetime"
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
      Index           =   17
      Left            =   6600
      TabIndex        =   8
      Top             =   2520
      Width           =   1815
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      DataField       =   "AR CUST Sales YTD"
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
      Index           =   18
      Left            =   2760
      TabIndex        =   7
      Top             =   2520
      Width           =   1815
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      DataField       =   "AR CUST Write Offs Last Year"
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
      Index           =   19
      Left            =   4680
      TabIndex        =   6
      Top             =   3600
      Width           =   1815
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      DataField       =   "AR CUST Write Offs Lifetime"
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
      Index           =   20
      Left            =   6600
      TabIndex        =   5
      Top             =   3600
      Width           =   1815
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      DataField       =   "AR CUST Write Offs YTD"
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
      Index           =   21
      Left            =   2760
      TabIndex        =   4
      Top             =   3600
      Width           =   1815
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      DataField       =   "AR CUST Average Days To Pay"
      DataSource      =   "adoPrimaryRS"
      Height          =   285
      Index           =   22
      Left            =   4680
      TabIndex        =   3
      Top             =   1680
      Width           =   1095
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   10485
      TabIndex        =   0
      Top             =   4170
      Width           =   10485
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   300
         Left            =   2160
         TabIndex        =   2
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Re&Calculate"
         Height          =   300
         Left            =   1080
         TabIndex        =   42
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   300
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Customer ID"
      DataSource      =   "adoPrimaryRS"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   41
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Customer Name"
      DataSource      =   "adoPrimaryRS"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   40
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      Caption         =   "0 - 30"
      DataSource      =   "adoPrimaryRS"
      Height          =   255
      Index           =   2
      Left            =   840
      TabIndex        =   39
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      Caption         =   "30 - 60"
      DataSource      =   "adoPrimaryRS"
      Height          =   255
      Index           =   3
      Left            =   2760
      TabIndex        =   38
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      Caption         =   "60 - 90"
      DataSource      =   "adoPrimaryRS"
      Height          =   255
      Index           =   4
      Left            =   4680
      TabIndex        =   37
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      Caption         =   "Over 90"
      DataSource      =   "adoPrimaryRS"
      Height          =   255
      Index           =   5
      Left            =   6600
      TabIndex        =   36
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      Caption         =   "Total"
      DataSource      =   "adoPrimaryRS"
      Height          =   255
      Index           =   6
      Left            =   8520
      TabIndex        =   35
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Highest Balance"
      DataSource      =   "adoPrimaryRS"
      Height          =   255
      Index           =   7
      Left            =   7200
      TabIndex        =   34
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Invoices"
      DataSource      =   "adoPrimaryRS"
      Height          =   255
      Index           =   10
      Left            =   1200
      TabIndex        =   33
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Payment Terms"
      DataSource      =   "adoPrimaryRS"
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   32
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Payments"
      DataSource      =   "adoPrimaryRS"
      Height          =   255
      Index           =   14
      Left            =   1080
      TabIndex        =   31
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Sales"
      DataSource      =   "adoPrimaryRS"
      Height          =   255
      Index           =   16
      Left            =   1200
      TabIndex        =   30
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Write Offs"
      DataSource      =   "adoPrimaryRS"
      Height          =   255
      Index           =   21
      Left            =   840
      TabIndex        =   29
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Average Days To Pay"
      DataSource      =   "adoPrimaryRS"
      Height          =   255
      Index           =   22
      Left            =   3000
      TabIndex        =   28
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      Caption         =   "YTD"
      DataSource      =   "adoPrimaryRS"
      Height          =   255
      Index           =   15
      Left            =   2760
      TabIndex        =   27
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      Caption         =   "Lifetime"
      DataSource      =   "adoPrimaryRS"
      Height          =   255
      Index           =   23
      Left            =   6600
      TabIndex        =   26
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      Caption         =   "Last Year"
      DataSource      =   "adoPrimaryRS"
      Height          =   255
      Index           =   24
      Left            =   4680
      TabIndex        =   25
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   10440
      Y1              =   4080
      Y2              =   4080
   End
End
Attribute VB_Name = "frm_AR_Cust_Financials"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ADOprimaryrs As ADODB.Recordset
Attribute ADOprimaryrs.VB_VarHelpID = -1
Dim db As ADODB.Connection
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean
'This form is for viewing purpose only no data are allowed to be changed manually.
'Logically only one Financial info per customer should be in the recordset

Private Sub Command1_Click()
  'Recalculate the values on the form
  'Recalc the Total AR Balance
  
  ADOprimaryrs.Fields("AR CUST Financial Total") = CCur(CDbl(Me.txtFields(5).Text) + CDbl(Me.txtFields(4).Text) + CDbl(Me.txtFields(3).Text) + CDbl(Me.txtFields(2).Text))
  
  'Recalc the YTD Fields
  'What is first day of year
  Dim DayOne As Variant
  DayOne = "1/01/" & Format(Now, "yyyy")
  
  Dim WriteOffSource As String
  WriteOffSource = " [AR Payment Header] LEFT JOIN [AR Payment Invoice Cross Reference] ON [AR Payment Header].[AR PAY ID] = [AR Payment Invoice Cross Reference].[AR CROSS Payment ID]"
  
  Sales = 0
  Returns = 0
  gCustomerID$ = Me.txtFields(0).Text
  
  Sales = SumRecord("[AR SALE Total]", "[AR Sales]", db, "[AR SALE Customer ID] = '" & gCustomerID$ & "' AND [AR SALE Posted YN] = TRUE AND [AR SALE Date] >= #" & DayOne & "# AND [AR SALE Document Type] in ('Invoice','Sales Memo','Beginning Balance','Finance Charge')")
  If IsNull(Sales) Then Sales = 0
      
  Returns = SumRecord("[AR SALE Total]", "[AR Sales]", db, "[AR SALE Customer ID] = '" & gCustomerID$ & "' AND [AR SALE Posted YN] = TRUE AND [AR SALE Date] >= #" & DayOne & "# AND [AR SALE Document Type] in ('Return','Credit Memo')")
  If IsNull(Returns) Then Returns = 0
  
  ADOprimaryrs.Fields("AR CUST Sales YTD") = Sales - Returns
  
  Payments = SumRecord("[AR PAY Amount]", "[AR Payment Header]", db, "[AR PAY Customer No] = '" & gCustomerID$ & "' AND [AR PAY Posted YN] = TRUE AND [AR PAY Transaction Date] >= #" & DayOne & "# AND [AR PAY NSF] = FALSE AND [AR PAY Type] <> 'Credit Memo' AND [AR PAY Type] <> 'Return'")
  If IsNull(Payments) Then
      ADOprimaryrs.Fields("AR CUST Payments YTD") = 0
  Else
      ADOprimaryrs.Fields("AR CUST Payments YTD") = Payments
  End If
 
  Invoices = CountRecord("[AR SALE Total]", "[AR Sales]", db, "[AR SALE Customer ID] = '" & gCustomerID$ & "' AND [AR SALE Posted YN] = TRUE AND [AR SALE Date] >= #" & DayOne & "# AND [AR SALE Document Type] in ('Invoice','Sales Memo','Beginning Balance','Finance Charge')")
  If IsNull(Invoices) Then
      ADOprimaryrs.Fields("AR CUST Invoices YTD") = 0
  Else
      ADOprimaryrs.Fields("AR CUST Invoices YTD") = Invoices
  End If
    
  WriteOff = SumRecord("[AR CROSS Write Off Amount]", WriteOffSource, db, "[AR PAY Customer No] = '" & gCustomerID$ & "' AND [AR PAY Transaction Date] >= #" & DayOne & "# AND [AR PAY Posted YN] = TRUE  AND [AR PAY NSF] = FALSE")
  If IsNull(WriteOff) Then
      ADOprimaryrs.Fields("AR CUST Write Offs YTD") = 0
  Else
      ADOprimaryrs.Fields("AR CUST Write Offs YTD") = WriteOff
  End If

  Sales = 0
  Returns = 0
  
  'Last Year
  Dim LastDay As Variant
  LastDay = DateAdd("y", -1, DayOne)
  DayOne = DateAdd("yyyy", -1, DayOne)
  
  Sales = SumRecord("[AR SALE Total]", "[AR Sales]", db, "[AR SALE Customer ID] = '" & gCustomerID$ & "' AND [AR SALE Posted YN] = TRUE AND [AR SALE Date] BETWEEN #" & DayOne & "# AND #" & LastDay & "#  AND [AR SALE Document Type] in ('Invoice','Sales Memo','Beginning Balance','Finance Charge')")
  If IsNull(Sales) Then Sales = 0
  
  Returns = SumRecord("[AR SALE Total]", "[AR Sales]", db, "[AR SALE Customer ID] = '" & gCustomerID$ & "' AND [AR SALE Posted YN] = TRUE AND [AR SALE Date] BETWEEN #" & DayOne & "# AND #" & LastDay & "#  AND [AR SALE Document Type] in ('Return','Credit Memo')")
    If IsNull(Returns) Then Returns = 0
  
  ADOprimaryrs.Fields("AR CUST Sales Last Year") = Sales - Returns
  
  Payments = SumRecord("[AR PAY Amount]", "[AR Payment Header]", db, "[AR PAY Customer No] = '" & gCustomerID$ & "' AND [AR PAY Posted YN] = TRUE AND [AR PAY Transaction Date] BETWEEN #" & DayOne & "# AND #" & LastDay & "# AND [AR PAY NSF] = FALSE AND [AR PAY Type] <> 'Credit Memo' AND [AR PAY Type] <> 'Return'")
  If IsNull(Payments) Then
    ADOprimaryrs.Fields("AR CUST Payments Last Year") = 0
  Else
    ADOprimaryrs.Fields("AR CUST Payments Last Year") = Payments
  End If
 
  Invoices = CountRecord("[AR SALE Total]", "[AR Sales]", db, "[AR SALE Customer ID] = '" & gCustomerID$ & "' AND [AR SALE Posted YN] = TRUE AND [AR SALE Date] BETWEEN #" & DayOne & "# AND #" & LastDay & "# AND [AR SALE Document Type] in ('Invoice','Sales Memo','Beginning Balance','Finance Charge')")
  If IsNull(Invoices) Then
    ADOprimaryrs.Fields("AR CUST Invoices Last Year") = 0
  Else
    ADOprimaryrs.Fields("AR CUST Invoices Last Year") = Invoices
 End If

  WriteOff = SumRecord("[AR CROSS Write Off Amount]", WriteOffSource, db, "[AR PAY Customer No] = '" & gCustomerID$ & "' AND [AR PAY Transaction Date] BETWEEN #" & DayOne & "# AND #" & LastDay & "# AND [AR PAY Posted YN] = TRUE  AND [AR PAY NSF] = FALSE")
  If IsNull(WriteOff) Then
    ADOprimaryrs.Fields("AR CUST Write Offs Last Year") = 0
  Else
    ADOprimaryrs.Fields("AR CUST Write Offs Last Year") = WriteOff
  End If

  Sales = 0
  Returns = 0

  'Lifetime
  Sales = SumRecord("[AR SALE Total]", "[AR Sales]", db, "[AR SALE Customer ID] = '" & gCustomerID$ & "' AND [AR SALE Posted YN] = TRUE AND [AR SALE Document Type] in ('Invoice','Sales Memo','Beginning Balance','Finance Charge')")
  If IsNull(Sales) Then Sales = 0
  
  Returns = SumRecord("[AR SALE Total]", "[AR Sales]", db, "[AR SALE Customer ID] = '" & gCustomerID$ & "' AND [AR SALE Posted YN] = TRUE AND [AR SALE Document Type] in ('Return','Credit Memo')")
  If IsNull(Returns) Then Returns = 0
  
  ADOprimaryrs.Fields("AR CUST Sales Lifetime") = Sales - Returns
  
  Payments = SumRecord("[AR PAY Amount]", "[AR Payment Header]", db, "[AR PAY Customer No] = '" & gCustomerID$ & "' AND [AR PAY Posted YN] = TRUE AND [AR PAY NSF] = FALSE AND [AR PAY Type] <> 'Credit Memo' AND [AR PAY Type] <> 'Return'")
  If IsNull(Payments) Then
    ADOprimaryrs.Fields("AR CUST Payments Lifetime") = 0
  Else
    ADOprimaryrs.Fields("AR CUST Payments Lifetime") = Payments
  End If
      
  Invoices = CountRecord("[AR SALE Total]", "[AR Sales]", db, "[AR SALE Customer ID] = '" & gCustomerID$ & "' AND [AR SALE Posted YN] = TRUE AND [AR SALE Document Type] in ('Invoice','Sales Memo','Beginning Balance','Finance Charge')")
    If IsNull(Invoices) Then
    ADOprimaryrs.Fields("AR CUST Invoices Lifetime") = 0
  Else
    ADOprimaryrs.Fields("AR CUST Invoices Lifetime") = Invoices
  End If
    
  WriteOff = SumRecord("[AR CROSS Write Off Amount]", WriteOffSource, db, "[AR PAY Customer No] = '" & gCustomerID$ & "' AND [AR PAY Posted YN] = TRUE  AND [AR PAY NSF] = FALSE")
  If IsNull(WriteOff) Then
    ADOprimaryrs.Fields("AR CUST Write Offs Lifetime") = 0
  Else
    ADOprimaryrs.Fields("AR CUST Write Offs Lifetime") = WriteOff
  End If
    
ADOprimaryrs.UpdateBatch adAffectAll
ADOprimaryrs.Requery
    
End Sub
Private Sub Form_Load()
On Error GoTo FormErr
    Set db = New Connection
    db.CursorLocation = adUseClient
    db.Open gblADOProvider
  
  Set ADOprimaryrs = New Recordset
  
  sql = "Select [AR CUST Customer ID],[AR CUST Name],[AR CUST Financial Period 1]," & _
    "[AR CUST Financial Period 2],[AR CUST Financial Period 3],[AR CUST Financial Period 4]," & _
    "[AR CUST Financial Total],[AR CUST Sales YTD],[AR CUST Sales Last Year]," & _
    "[AR CUST Sales Lifetime],[AR CUST Payments YTD],[AR CUST Payments Last Year]," & _
    "[AR CUST Payments Lifetime],[AR CUST Write Offs YTD],[AR CUST Write Offs Last Year]," & _
    "[AR CUST Write Offs Lifetime],[AR CUST Invoices YTD],[AR CUST Invoices Last Year]," & _
    "[AR CUST Invoices Lifetime],[AR CUST Highest Balance],[AR CUST Average Days To Pay]," & _
    "[AR CUST Payment Terms] From [AR Customer]" '====
  SQLW = " Where [AR CUST Customer ID] = '" & frm_AR_Customer.txtFields(0).Text & "'"
  SQLState = sql & SQLW
    
  If IsNull(frm_AR_Customer.txtFields(0).Text) Then
    ADOprimaryrs.Open sql, db, adOpenStatic, adLockOptimistic
  Else
    ADOprimaryrs.Open SQLState, db, adOpenStatic, adLockOptimistic, adCmdText
  End If
  
  Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtFields
    Set oText.DataSource = ADOprimaryrs
    oText.Locked = True
  Next
  GetTextColor Me
  mbDataChanged = False
Exit Sub
FormErr:
  MsgBox Err.Description
  Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
  'On Error Resume Next
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If mbEditFlag Or mbAddNewFlag Then Exit Sub

  If KeyCode = vbKeyEscape Then cmdClose_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo FormErr
    'updates the checklist Vendors
  Screen.MousePointer = vbHourglass
      If ADOprimaryrs.RecordCount > 0 Then
        If ADOprimaryrs.EditMode <> 0 Then
          ADOprimaryrs.CancelUpdate
        End If
      End If
      ADOprimaryrs.Close
      Set ADOprimaryrs = Nothing
      db.Close
      Set db = Nothing
  Screen.MousePointer = vbDefault
  Set frm_AR_Cust_Financials = Nothing
  Exit Sub
FormErr:
  MsgBox Err.Description
  Screen.MousePointer = vbDefault
End Sub

Private Sub adoPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This will display the current record position for this recordset
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

Private Sub cmdRefresh_Click()
  'This is only needed for multi user apps
  'On Error GoTo RefreshErr
  ADOprimaryrs.Requery
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub SetButtons(bVal As Boolean)
  cmdAdd.Visible = bVal
  cmdUpdate.Visible = bVal
  cmdDelete.Visible = bVal
  cmdClose.Visible = bVal
  cmdRefresh.Visible = bVal
  cmdNext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdLast.Enabled = bVal
  cmdPrevious.Enabled = bVal
End Sub

