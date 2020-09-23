VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form AllLookup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LookUP"
   ClientHeight    =   3765
   ClientLeft      =   5970
   ClientTop       =   3540
   ClientWidth     =   4140
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   4140
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   330
      Width           =   2535
   End
   Begin VB.CommandButton CmdSelect 
      Default         =   -1  'True
      Height          =   615
      Left            =   3480
      Picture         =   "frm_lookup.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Select"
      Top             =   50
      Width           =   615
   End
   Begin MSDataGridLib.DataGrid grdAllLookup 
      Height          =   3015
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   4130
      _ExtentX        =   7276
      _ExtentY        =   5318
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   4
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdSearch 
      Height          =   615
      Left            =   2760
      Picture         =   "frm_lookup.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Search"
      Top             =   50
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Search Criteria"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "AllLookup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents ADOprimaryrs As ADODB.Recordset
Attribute ADOprimaryrs.VB_VarHelpID = -1
'Dim mbDataChanged As Boolean

Dim TempStr As String
Dim GridHeader As String

Dim CallFormType As Integer
Dim MovetoCommand As Boolean
Dim dbTemp As ADODB.Connection
Dim WhichField As String


Private Function CheckOwner() As Boolean
CheckOwner = True
     If ADOprimaryrs![AR CUST SalesPerson] <> AppLoginName Then
        Dim Response As Integer
            Response = MsgBox("This account is belong to " & ADOprimaryrs![AR CUST SalesPerson] & _
            "." & vbCr & "Would you like to continue?", vbYesNo, "Information")
            If Response = vbNo Then
                CheckOwner = False
            Else
                CheckOwner = True
            End If
     End If
  
End Function

Private Sub cmdSearch_Click()
If MovetoCommand = True Then
    MsgBox "We can searh on browsing appearance"
    Exit Sub
End If
If Text1 = "" Then Exit Sub
    If SearchData(False) = False Then
        MsgBox Label1 & " " & Text1 & " is not existed.", vbInformation, "Information"
    End If
End Sub

Private Function SearchData(ConfirmPost As Boolean) As Boolean
On Error GoTo NOTFOUND
Dim SearchField As String
Dim Selstr As String

    Selstr = Text1.Text
    If ConfirmPost = True Or WhichField = "" Then
        SearchField = grdAllLookup.Columns(0).DataField
    Else
        'SearchField = grdAllLookup.Columns(SelCol).DataField
        SearchField = WhichField
    End If

    grdAllLookup.SetFocus
    ADOprimaryrs.MoveFirst
    If ADOprimaryrs("" & SearchField & "").Type = 202 Then
        ADOprimaryrs.Find "[" & SearchField & "]='" & Selstr & "'"
    Else
        ADOprimaryrs.Find "[" & SearchField & "]=" & Selstr
    End If
    SendKeys ("{LEFT}")
    SearchData = True
    
    If ADOprimaryrs.EOF = True Then
NOTFOUND:
        SearchData = False
    End If
End Function

Private Sub CmdSelect_Click()
Dim i As Integer
Dim GridCol
Dim dgridSource
Dim dgridTarget

If MovetoCommand = True Then
    GoTo Out_Of_Here
End If

If Text1 = "" Then
   MsgBox "Please Make A Selection", vbCritical, "Error"
   Exit Sub
End If

If SearchData(True) = False Then
   MsgBox "The choosen Text is not within the " & grdAllLookup.ColumnHeaders & " range", vbCritical, "Error"
   Exit Sub
Else
    Text1 = grdAllLookup.Columns(0).Value
End If
'Exit Sub
'Me.Hide

     'fMainForm.ActiveForm.grdDataGrid.AllowUpdate = True
'1000 - 1099 --> frm_AR_Sales_Entry
'1100 - 1199 --> Main_Options
'1200        --> frm_Sales_Posting_Payments

Select Case CallFormType
    'Yusri's Programming
    Case 1  ' updates  field GL Sales Account in form frm_AR_Customer
        frm_AR_Customer.txtFields(18).Text = grdAllLookup.Columns(0).Value
    Case 2  ' updates field GL Account Number in form frm_LIST_Credit_Card
        frm_LIST_Credit_Cards.txtFields(4).Text = grdAllLookup.Columns(0).Value
    Case 3  'updates field Inventory Adjustment Account in form _
                frm_SYS_Setup_Inventory
        frm_SYS_Setup_Inventory.txtFields(0).Text = grdAllLookup.Columns(0).Value
    Case 4  'updates field Inventory Production Account in tab Inventory in _
                form frm_SYS_Setup_Inventory
        frm_SYS_Setup_Inventory.txtFields(2).Text = grdAllLookup.Columns(0).Value
    Case 5 'updates field Vendor ID in tab Items in frm_SYS_Setup_Items
        frm_SYS_Setup_Items.txtFields(6).Text = grdAllLookup.Columns(0).Value
    Case 6  'updates field GL Sales Account in frm_SYS_Setup_Items
        frm_SYS_Setup_Items.txtFields(8).Text = grdAllLookup.Columns(0).Value
    Case 7 'updates field GL Inventory Account in tab Items in frm_SYS_Setup_Items
        frm_SYS_Setup_Items.txtFields(9).Text = grdAllLookup.Columns(0).Value
    Case 8 'updates field GL Cost Of Sales in frm_SYS_Setup_Items
        fMainForm.ActiveForm.txtFields(10).Text = grdAllLookup.Columns(0).Value
    Case 9 'updates field GL Number in frm frm_SYS_Setup_GL
        fMainForm.ActiveForm.txtFields(1).Text = grdAllLookup.Columns(0).Value
    Case 10 ' updates field GL Bank Interest Earned Accountin frm_SYS_Setup_Banks
        frm_SYS_Setup_Banks.txtFields(0).Text = grdAllLookup.Columns(0).Value
    Case 11 'updates field GL Bank Misc Deposit Account in frm_SYS_Setup_Banks
        frm_SYS_Setup_Banks.txtFields(1).Text = grdAllLookup.Columns(0).Value
    Case 12 'updates field GL Bank Misc Withdrawal Account in frm_SYS_Setup_Banks
        frm_SYS_Setup_Banks.txtFields(2).Text = grdAllLookup.Columns(0).Value
    Case 13 'updates field GL Bank Other Charges Account in frm_SYS_Setup_Banks
        frm_SYS_Setup_Banks.txtFields(3).Text = grdAllLookup.Columns(0).Value
    Case 14 'updates field GL Bank Service Charges Account in frm_SYS_Setup_Banks
        frm_SYS_Setup_Banks.txtFields(4).Text = grdAllLookup.Columns(0).Value
    Case 15 'updates field GL Bank Finance Charges Account in frm_SYS_Setup_Banks
        frm_SYS_Setup_Banks.txtFields(5).Text = grdAllLookup.Columns(0).Value
    Case 16 'updates field GL Purchase Account in frm_SYS_Setup_Purchases
        frm_SYS_Setup_Purchases.txtFields(0).Text = grdAllLookup.Columns(0).Value
    Case 17 'updates field GL Purchase AP Account in frm_SYS_Setup_Purchases
        frm_SYS_Setup_Purchases.txtFields(1).Text = grdAllLookup.Columns(0).Value
    Case 18 'updates field GL Purchase Inventory Account in frm_SYS_Setup_Purchases
        frm_SYS_Setup_Purchases.txtFields(2).Text = grdAllLookup.Columns(0).Value
    Case 19 'updates field GL Purchase Cash Account in frm_SYS_Setup_Purchases
        frm_SYS_Setup_Purchases.txtFields(3).Text = grdAllLookup.Columns(0).Value
    Case 20 'updates field GL Purchase Pre Paid Account in frm_SYS_Setup_Purchases
        frm_SYS_Setup_Purchases.txtFields(4).Text = grdAllLookup.Columns(0).Value
    Case 21 'updates field GL Purchase Freight Account in frm_SYS_Setup_Purchases
        frm_SYS_Setup_Purchases.txtFields(5).Text = grdAllLookup.Columns(0).Value
    Case 22 'updates field GL Purchase Discount Account in frm_SYS_Setup_Purchases
        frm_SYS_Setup_Purchases.txtFields(6).Text = grdAllLookup.Columns(0).Value
    Case 23 'updates field GL Purchase Misc Account in frm_SYS_Setup_Purchases
        frm_SYS_Setup_Purchases.txtFields(7).Text = grdAllLookup.Columns(0).Value
    Case 24 'updates field GL Purchase Write Off Account in frm_SYS_Setup_Purchases
        frm_SYS_Setup_Purchases.txtFields(8).Text = grdAllLookup.Columns(0).Value
    Case 25 'updates field GL Sales Account Default in frm_SYS_Setup_Sales
        frm_SYS_Setup_Sales.txtFields(0).Text = grdAllLookup.Columns(0).Value
    Case 26 'updates field GL Sales AR Account in frm_SYS_Setup_Sales
        frm_SYS_Setup_Sales.txtFields(1).Text = grdAllLookup.Columns(0).Value
    Case 27 'updates field GL Sales Cash Account in frm_SYS_Setup_Sales
        frm_SYS_Setup_Sales.txtFields(2).Text = grdAllLookup.Columns(0).Value
    Case 28 'updates field GL Sales Cogs Account in frm_SYS_Setup_Sales
        frm_SYS_Setup_Sales.txtFields(3).Text = grdAllLookup.Columns(0).Value
    Case 29 'updates field GL Sales Discount Account in frm_SYS_Setup_Sales
        frm_SYS_Setup_Sales.txtFields(4).Text = grdAllLookup.Columns(0).Value
    Case 30 'updates field GL Freight Account in frm_SYS_Setup_Sales
        frm_SYS_Setup_Sales.txtFields(5).Text = grdAllLookup.Columns(0).Value
    Case 31 'updates field GL Sales Inventory Account in frm_SYS_Setup_Sales
        frm_SYS_Setup_Sales.txtFields(6).Text = grdAllLookup.Columns(0).Value
    Case 32 ' updates field GL Sales Misc Account in frm_SYS_Setup_Sales
        frm_SYS_Setup_Sales.txtFields(7).Text = grdAllLookup.Columns(0).Value
    Case 33 'updates field GL Sales Sales Account in frm_SYS_Setup_Sales
        frm_SYS_Setup_Sales.txtFields(8).Text = grdAllLookup.Columns(0).Value
    Case 34 'updates field GL Sales Tax Account in frm_SYS_Setup_Sales
        frm_SYS_Setup_Sales.txtFields(9).Text = grdAllLookup.Columns(0).Value
    Case 35 'updates field GL Sales Return in frm_SYS_Setup_Sales
        frm_SYS_Setup_Sales.txtFields(10).Text = grdAllLookup.Columns(0).Value
    Case 36 'updates field GL Sales Write Off in frm_SYS_Setup_Sales
        frm_SYS_Setup_Sales.txtFields(0).Text = grdAllLookup.Columns(0).Value
    Case 37 'updates field GL Sales Write Off in frm_Cust_Projects
        fMainForm.ActiveForm.txtFields(6).Text = grdAllLookup.Columns(0).Value
    Case 38 'updates field Default GL in frm_AP_Vendor
        fMainForm.ActiveForm.txtFields(30).Text = grdAllLookup.Columns(0).Value
        
    'Razi's Programming
    ' to update the customer data in the frm_ar_Quote/Order/sales_entry
    Case 1000
        If CheckOwner = False Then Exit Sub
        For i = 0 To grdAllLookup.Columns.count - 1
            Set GridCol = grdAllLookup.Columns(i)
                fMainForm.ActiveForm.txtFieldsCust(i) = GridCol
        Next
    ' to update the shipping location in the frm_ar_Quote/Order/sales_entry
    Case 1001  'frm_ar_sales_entry
        
        For i = 0 To grdAllLookup.Columns.count - 1
            Set GridCol = grdAllLookup.Columns(i)
                fMainForm.ActiveForm.txtFieldsShip(i) = GridCol
        Next
    Case 1002  'Modify the GL Account frm_ar_Quote/Order/sales_entry
                fMainForm.ActiveForm.grdDataGrid.Columns(9).Text = Text1
    Case 1003  'Add the project frm_ar_Quote/Order/sales_entry
                fMainForm.ActiveForm.grdDataGrid.Columns(10).Text = Text1
    Case 1004  'frm_ar_Quote_entry for datagrid
            'Dim dgridSource
            'Dim dgridTarget
                        
            Set dgridSource = grdAllLookup.Columns
            Set dgridTarget = fMainForm.ActiveForm.grdDataGrid.Columns
                    
            'update the sub total on frm_ar_Quote_entry
            dgridTarget(2).Text = 0
            'dgridTarget(3).Text = 0
            dgridTarget(5).Text = 0
            dgridTarget(6).Text = 0
            'dgridTarget(8).Text = 0
                        
            dgridTarget(0).Text = dgridSource(0).Value
            dgridTarget(1).Text = dgridSource(1).Value
            dgridTarget(3).Text = dgridSource(2).Value
            dgridTarget(4).Text = dgridSource(3).Value
            dgridTarget(8).Text = dgridSource(8).Value
            dgridTarget(9).Text = dgridSource(4).Value
                
                'If dgridSource(7).Value = -1 Then
                '   dgridTarget(7).Text = "No"
                'Else
                '   dgridTarget(7).Text = "Yes"
                'End If
                    
                'If fMainForm.ActiveForm.grdOnAddNew = True Then
                '    Me.MousePointer = vbHourglass
                '    fMainForm.ActiveForm.grdDataGrid.Row = fMainForm.ActiveForm.grdDataGrid.Row + 1
                '    fMainForm.ActiveForm.AddNewROW_to_grd
                '    Me.MousePointer = vbNormal
                'End If
    ' for sales setup -cmdSales
    Case 1005  'frm_ar_Order_entry for datagrid
            'Dim dgridSource
            'Dim dgridTarget
                        
            Set dgridSource = grdAllLookup.Columns
            Set dgridTarget = fMainForm.ActiveForm.grdDataGrid.Columns
                    
            'update the sub total on frm_order_entry
            dgridTarget(2).Text = 0
            dgridTarget(3).Text = 0
            dgridTarget(6).Text = 0
            dgridTarget(8).Text = 0
                        
            dgridTarget(0).Text = dgridSource(0).Value
            dgridTarget(1).Text = dgridSource(1).Value
            dgridTarget(4).Text = dgridSource(2).Value
            dgridTarget(5).Text = dgridSource(3).Value
            'dgridTarget(8).Text = dgridSource(8).Value
            dgridTarget(9).Text = dgridSource(4).Value
                
                'If fMainForm.ActiveForm.grdOnAddNew = True Then
                '    Me.MousePointer = vbHourglass
                '    fMainForm.ActiveForm.grdDataGrid.Row = fMainForm.ActiveForm.grdDataGrid.Row + 1
                '    fMainForm.ActiveForm.AddNewROW_to_grd
                '    Me.MousePointer = vbNormal
                'End If
    Case 1006  'frm_ar_Sales_entry for datagrid
            'Dim dgridSource3
            'Dim dgridTarget4
                        
            Set dgridSource = grdAllLookup.Columns
            Set dgridTarget = fMainForm.ActiveForm.grdDataGrid.Columns
                    
            'update the sub total on frm_order_entry
            dgridTarget(2).Text = 0
            'dgridTarget(3).Text = 0
            dgridTarget(6).Text = 0
            dgridTarget(8).Text = 0
                        
            dgridTarget(0).Text = dgridSource(0).Value
            dgridTarget(1).Text = dgridSource(1).Value
            dgridTarget(3).Text = dgridSource(2).Value
            dgridTarget(4).Text = dgridSource(8).Value
            dgridTarget(5).Text = dgridSource(3).Value
            dgridTarget(9).Text = dgridSource(4).Value
                
                'If fMainForm.ActiveForm.grdOnAddNew = True Then
                '    Me.MousePointer = vbHourglass
                '    fMainForm.ActiveForm.grdDataGrid.Row = fMainForm.ActiveForm.grdDataGrid.Row + 1
                '    fMainForm.ActiveForm.AddNewROW_to_grd
                '    Me.MousePointer = vbNormal
                'End If
    Case 1008   'AR_Sales_Memo_Entry
            Set dgridSource = grdAllLookup.Columns
            Set dgridTarget = fMainForm.ActiveForm.grdDataGrid.Columns
            dgridTarget(2).Text = "$0.00"
            dgridTarget(0).Text = dgridSource(0).Value
            dgridTarget(1).Text = dgridSource(1).Value
    
    Case 1010  'Add the project frm_ar_Quote/Order/sales_entry
                fMainForm.ActiveForm.grdDataGrid.Columns(10).Text = Text1
    Case 1100  'frm_SYS_Setup_Financial
            i = Int(fMainForm.ActiveForm.lblLookupNumbers.Caption)
            fMainForm.ActiveForm.txtSales(i).Text = Text1
            fMainForm.ActiveForm.lblGLAccounts(i) = grdAllLookup.Columns(1).Value
            fMainForm.ActiveForm.lblGLAccounts(i).Visible = True
    Case 1200   'frm_AR_Sales_Posting_Payments
            fMainForm.ActiveForm.txtFields(0) = Text1
    Case 1210   'frm_AR_Cash_Receipts for txtFields(0)
            fMainForm.ActiveForm.txtFields(0) = Text1
            fMainForm.ActiveForm.lblcashReceiptsTrue = grdAllLookup.Columns(1)
    Case 1215   'frm_AR_Cash_Receipts for txtFields(0)
            fMainForm.ActiveForm.txtFields(5) = Text1
    Case 1216   'frm_AR_Cash_Receipts for txtFields(1)
            fMainForm.ActiveForm.txtFields(1) = Text1
    Case 1220   'frm_AR_Order_Entry for txtfields(35)
            fMainForm.ActiveForm.txtFields(35) = Text1
    Case 1230   'frm_AR_Order_Entry for txtfields(35)
            fMainForm.ActiveForm.txtFields(35) = Text1
            fMainForm.ActiveForm.txtFields(34) = grdAllLookup.Columns(2)
    Case 1300   'frm_AP_Purchase_Entry for txtFieldsVendor
          For i = 0 To fMainForm.ActiveForm.txtFieldsVendor.UBound
            fMainForm.ActiveForm.txtFieldsVendor(i) = grdAllLookup.Columns(i)
          Next
    Case 1302   'frm_AP_Purchase_Entry for grddatagrid.columns(6)
                fMainForm.ActiveForm.grdDataGrid.Columns(6).Text = Text1
    Case 1304   'frm_AP_Purchase_Entry for grddatagrid.columns(7)
                fMainForm.ActiveForm.grdDataGrid.Columns(7).Text = Text1
    Case 1305   'frm_AR_Credit_Entry for grddatagrid.columns(3)
                fMainForm.ActiveForm.grdDataGrid.Columns(3).Text = Text1
    Case 1306   'frm_AP_Purchase_Entry for grddatagrid.columns(0) and more
                        
            Set dgridSource = grdAllLookup.Columns
            'fMainForm.ActiveForm.grdDataGrid.SetFocus
            Set dgridTarget = fMainForm.ActiveForm.grdDataGrid.Columns
                    
            'update the sub total on frm_order_entry
            dgridTarget(2).Text = 0
            dgridTarget(5).Text = 0
                        
            dgridTarget(0).Text = dgridSource(0).Value
            dgridTarget(1).Text = dgridSource(1).Value
            dgridTarget(3).Text = dgridSource(2).Value
            dgridTarget(4).Text = dgridSource(8).Value
            dgridTarget(6).Text = dgridSource(4).Value
                    
                'If fMainForm.ActiveForm.grdOnAddNew = True Then
                '    Me.MousePointer = vbHourglass
                    'fMainForm.ActiveForm.grdDataGrid.Row = fMainForm.ActiveForm.grdDataGrid.Row + 1
                '    fMainForm.ActiveForm.AddNewROW_to_grd
                '    Me.MousePointer = vbNormal
                'End If
    Case 1400   'frm_AP_Voucher_Entry for grddatagrid.columns(0) and more
                        
            Set dgridSource = grdAllLookup.Columns
            'fMainForm.ActiveForm.grdDataGrid.SetFocus
            Set dgridTarget = fMainForm.ActiveForm.grdDataGrid.Columns
                    
            'update the sub total on frm_order_entry
            dgridTarget(2).Text = "$0.00"
                        
            dgridTarget(0).Text = dgridSource(0).Value  'item id
            dgridTarget(1).Text = dgridSource(1).Value  'description
                    
    Case 1350   'frm_AP_Purchase_Entry for grddatagrid.columns(0) and more
                        
            Set dgridSource = grdAllLookup.Columns
            'fMainForm.ActiveForm.grdDataGrid.SetFocus
            Set dgridTarget = fMainForm.ActiveForm.grdDataGrid.Columns
                    
            'update the sub total on frm_order_entry
            dgridTarget(2).Text = 0
            dgridTarget(5).Text = 0
            dgridTarget(6).Text = FormatDate(Now)
            dgridTarget(7).Text = "N/A"
                        
            dgridTarget(0).Text = dgridSource(0).Value  'item id
            dgridTarget(1).Text = dgridSource(1).Value  'description
            dgridTarget(4).Text = dgridSource(2).Value  'Unit
            dgridTarget(5).Text = dgridSource(8).Value  'Cost price
            dgridTarget(9).Text = dgridSource(4).Value  'Sales Account
    Case 1401   'frm_INV_Adjust for grddatagrid.columns(0) and more
                        
            Set dgridSource = grdAllLookup.Columns
            'fMainForm.ActiveForm.grdDataGrid.SetFocus
            Set dgridTarget = fMainForm.ActiveForm.grdDataGrid.Columns
            
            dgridTarget(4).Text = 0
            
            dgridTarget(0).Text = dgridSource(0).Value  'item id
            dgridTarget(1).Text = dgridSource(1).Value  'description
            dgridTarget(2).Text = dgridSource(3).Value  'inventory Acct
            'dgridTarget(3).Text = LookRecord("[GL COA Account Name]", "[GL Chart Of Accounts]", "[GL COA Account No] = '" & dgridSource(3).Value & "'")  'COA Name
            dgridTarget(3).Text = dgridSource(4).Value  'Qty on hand
            If fMainForm.ActiveForm.cbfields(6).Text = "Increase" Then
              dgridTarget(5).Text = IIf(IsNull(dgridSource(6).Value), 0, dgridSource(6).Value)
            Else
              dgridTarget(5).Text = IIf(IsNull(dgridSource(7).Value), 0, dgridSource(7).Value)
            End If
    Case 1403  'Add the project frm_inv_adjust
                fMainForm.ActiveForm.grdDataGrid.Columns(7).Text = Text1
    Case 1410  'Add the project frm_inv_adjust
                fMainForm.ActiveForm.grdDataGrid.Columns(2).Text = Text1
                'Debug.Print LookRecord("[GL COA Account Name]", "[GL Chart Of Accounts]", "[GL COA Account No] = '" & fMainForm.ActiveForm.grdDataGrid.Columns(2).Text & "'") 'COA Name
                'fMainForm.ActiveForm.grdDataGrid.Columns(3).Text = LookRecord("[GL COA Account Name]", "[GL Chart Of Accounts]", "[GL COA Account No] = '" & fMainForm.ActiveForm.grdDataGrid.Columns(2).Text & "'") 'COA Name
    Case 1420 'frm_INV_Production for grddatagrid.columns(0) and more
                        
            Set dgridSource = grdAllLookup.Columns
            'fMainForm.ActiveForm.grdDataGrid.SetFocus
            Set dgridTarget = fMainForm.ActiveForm.grdDataGrid.Columns
            
            dgridTarget(2).Text = 0
            dgridTarget(4).Text = "$0.00"
            
            dgridTarget(0).Text = dgridSource(0).Value  'item id
            dgridTarget(1).Text = dgridSource(1).Value  'description
            dgridTarget(3).Text = dgridSource(6).Value  'Last Cost
            'dgridTarget(3).Text = dgridSource(4).Value  'Qty on hand
            'If fMainForm.ActiveForm.cbfields(6).Text = "Increase" Then
            '  dgridTarget(5).Text = IIf(IsNull(dgridSource(6).Value), 0, dgridSource(6).Value)
            'Else
            '  dgridTarget(5).Text = IIf(IsNull(dgridSource(7).Value), 0, dgridSource(7).Value)
            'End If
    Case 1423  'Add the project frm_inv_adjust
                fMainForm.ActiveForm.grdDataGrid.Columns(5).Text = Text1
    Case 1450   'Add the COA frm_GL_Entry
                fMainForm.ActiveForm.grdDataGrid.Columns(0).Text = Text1
    Case 1453  'Add the project frm_GL_Entry
                fMainForm.ActiveForm.grdDataGrid.Columns(3).Text = Text1
    Case 1460  'Add the project frm_Bank_Transaction
                fMainForm.ActiveForm.txtFields(4).Text = Text1
    Case 1465  'Add the project frm_Bank_Transaction
                fMainForm.ActiveForm.txtFields(12).Text = Text1
    Case 1480  'Add the project frm_Bank_Reconciliation
                fMainForm.ActiveForm.txtFields(2).Text = Text1
                fMainForm.ActiveForm.lblLabels(6).Caption = grdAllLookup.Columns(1).Value
    Case 1490  'Add the project frm_Bank_Reconciliation
                fMainForm.ActiveForm.txtFields(12).Text = Text1
                fMainForm.ActiveForm.lblLabels(12).Caption = grdAllLookup.Columns(1).Value
    Case 1499  'Add the project frm_Bank_Reconciliation
                fMainForm.ActiveForm.txtFields(1).Text = Text1
                fMainForm.ActiveForm.lblLabels(23).Caption = grdAllLookup.Columns(1).Value
    Case 1510  'Add the project frm_Pay_Employees
                fMainForm.ActiveForm.txtFields.Text = Text1
                fMainForm.ActiveForm.lblAccounts.Caption = grdAllLookup.Columns(1).Value
    Case 1520  'Add the project frm_Pay_Voids
                fMainForm.ActiveForm.txtFields(0).Text = Text1
                fMainForm.ActiveForm.lblAccts.Caption = grdAllLookup.Columns(1).Value
    Case 1530  'Add the project frm_Pay_Voids
                fMainForm.ActiveForm.txtFields(1).Text = Text1
                fMainForm.ActiveForm.lblEmp.Caption = grdAllLookup.Columns(1).Value
    Case 1540  'Add the project frm_Pay_Voids
                fMainForm.ActiveForm.txtFields(2).Text = Text1
                fMainForm.ActiveForm.txtFields(3).Text = FormatDate(grdAllLookup.Columns(1).Value)
                fMainForm.ActiveForm.txtFields(4).Text = FormatCurr(grdAllLookup.Columns(2).Value)
                fMainForm.ActiveForm.txtFields(5).Text = grdAllLookup.Columns(4).Value
    Case 1550  'Add the project frm_Pay_Voids
                fMainForm.ActiveForm.txtPyrllItems(0).Text = Text1
    Case 1560  'Add the project frm_Pay_Employees
                fMainForm.ActiveForm.txtPyrllItems(8).Text = Text1
                fMainForm.ActiveForm.lblAccts.Caption = grdAllLookup.Columns(1).Value
    Case 1570  'Add the project frm_Pay_Employees
                fMainForm.ActiveForm.txtPyrllItems(11).Text = Text1
                fMainForm.ActiveForm.lblDebit.Caption = grdAllLookup.Columns(1).Value
    Case 1580  'Add the project frm_Pay_Employees
                fMainForm.ActiveForm.txtPyrllItems(12).Text = Text1
                fMainForm.ActiveForm.lblCredit.Caption = grdAllLookup.Columns(1).Value
    Case 1590  'Add the project frm_Pay_Employees
                fMainForm.ActiveForm.txtCommission(0).Text = Text1
    Case 1600  'Add the project frm_SYS_Setup_Payroll
                fMainForm.ActiveForm.txtFieldsTemp.Text = Text1
                fMainForm.ActiveForm.lblAcctTemp.Caption = grdAllLookup.Columns(1).Value
    Case 1610  'Add the project frm_SYS_Setup_Employees
                fMainForm.ActiveForm.txtPyrllItems(0).Text = Text1
    Case 1650  'Add the project frm_SYS_Setup_Accounting_Preferences
                fMainForm.ActiveForm.txtFieldsTemp.Text = Text1
                fMainForm.ActiveForm.lblAcct(0).Caption = grdAllLookup.Columns(1).Value
    Case 1660  'Add the project frm_SYS_Tax_Group
                Dim ColCount As Integer
                Dim SameData As Boolean
                'MsgBox fMainForm.ActiveForm.grdDataGrid.Row
                ColCount = fMainForm.ActiveForm.grdDataGrid.Row
                SameData = True
                If ColCount > 0 Then
                For i = 0 To ColCount
                    fMainForm.ActiveForm.grdDataGrid.Row = i
                    'MsgBox fMainForm.ActiveForm.grdDataGrid.Columns(0).Text
                    If Text1 = fMainForm.ActiveForm.grdDataGrid.Columns(0).Text Then
                        SameData = False
                        MsgBox "Can't insert the same data", vbCritical, "Error"
                    End If
                Next
                Else
                    SameData = True
                End If
                fMainForm.ActiveForm.grdDataGrid.Row = ColCount
                If SameData = True And frm_SYS_Setup_Tax_Group.grdOnAddNew = True Then
                    fMainForm.ActiveForm.grdDataGrid.Columns(0).Text = Text1
                    fMainForm.ActiveForm.grdDataGrid.Row = ColCount + 1
                End If
    Case 1700   'check detail --- frm_Pay_Employees
                fMainForm.ActiveForm.txtCheckDetail(0) = Text1
    Case 1750   'frm_SYS_Setup_Items for DataGrid1.columns(0) and more
                        
            Set dgridSource = grdAllLookup.Columns
            Set dgridTarget = fMainForm.ActiveForm.DataGrid1.Columns
            
            dgridTarget(3).Text = 0
                        
            dgridTarget(2).Text = dgridSource(0).Value
            dgridTarget(4).Text = dgridSource(2).Value  'Unit
            dgridTarget(5).Text = dgridSource(8).Value  'Unit
End Select


Out_Of_Here:
    Unload Me  ' Unload the current form
End Sub

Private Sub Form_Load()
If dbTemp Is Nothing Then
    Set dbTemp = New ADODB.Connection
    dbTemp.CursorLocation = adUseClient
    dbTemp.Open gblADOProvider
End If
    grdAllLookup.ClearFields
End Sub

Public Sub GetWhichTable(FormType As Integer, SQLstatement As String, GridName As String, grdDataGridHeader As String, Optional db As ADODB.Connection)

MovetoCommand = False 'this is for selection task
'put formtype in public variable
    CallFormType = FormType
'give the caption name to the datagrid
    Caption = "Lookup - " & GridName
    grdAllLookup.Caption = GridName
'add or remove column to the datagrid and name for the header
Dim TotalChar, i, j As Integer
Dim storedString As String

TempStr = SQLstatement
GridHeader = grdDataGridHeader
'Important
'A DataGrid object can contain only 32767 columns, as column indices are stored
'in integers.

'If you have previously deleted a column using the Remove method, after adding
'new columns, you may need to refresh the display with the Rebind and Refresh
'methods. This instructs the DataGrid control to rebuild its internal column
'layout matrix to correctly reflect the true status of the control.
'---remove unnecessary column
    TotalChar = Len(grdDataGridHeader)
        If grdAllLookup.Columns.count > 0 Then
            i = grdAllLookup.Columns.count - 1
            While i <> 0
                grdAllLookup.Columns.Remove (i)
                i = i - 1
            Wend
        End If
            
  'database declaration
  If db Is Nothing Then
    Set db = New ADODB.Connection
    db.CursorLocation = adUseClient
    db.Open gblADOProvider
  End If

  Set ADOprimaryrs = New ADODB.Recordset
  
  'getting the datagrid column header
  
  ADOprimaryrs.Open SQLstatement, db, adOpenStatic, adLockReadOnly, adCmdText
  
  'checking for data existance
  If ADOprimaryrs.RecordCount = 0 Then
    MsgBox "No data to publish.", vbInformation, "I'm sorry"
    GoTo NoDataToPublish
  End If
  
  Set grdAllLookup.DataSource = ADOprimaryrs
  
  'mbDataChanged = False
     
     grdAllLookup.ClearFields
     grdAllLookup.ReBind
     grdAllLookup.Refresh
  
  Dim grdWidth As Integer
  grdWidth = 0
  'Change The datagrid columnheader name
    For i = 1 To TotalChar
         
         If Mid(grdDataGridHeader, i, 2) = "//" Then
            grdDataGridHeader = Right(grdDataGridHeader, TotalChar - i - 1)
            TotalChar = Len(grdDataGridHeader)
            grdAllLookup.Columns(j).Caption = storedString
            
            If Len(Trim(ADOprimaryrs("" & grdAllLookup.Columns(j).DataField & ""))) < 8 Then
                grdAllLookup.Columns(j).Width = 800
            ElseIf Len(Trim(ADOprimaryrs("" & grdAllLookup.Columns(j).DataField & ""))) < 30 Then
                grdAllLookup.Columns(j).Width = 2000
            Else
                grdAllLookup.Columns(j).Width = 2000
            End If
            grdWidth = grdWidth + grdAllLookup.Columns(j).Width
            
            j = j + 1
            storedString = ""
            i = 1
            If TotalChar = 0 Then Exit For
         End If
         storedString = Left(grdDataGridHeader, i)
    Next
    grdAllLookup.Columns(j).Caption = storedString
    If Len(Trim(ADOprimaryrs("" & grdAllLookup.Columns(j).DataField & ""))) < 8 Then
        grdAllLookup.Columns(j).Width = 800
    ElseIf Len(Trim(ADOprimaryrs("" & grdAllLookup.Columns(j).DataField & ""))) < 30 Then
        grdAllLookup.Columns(j).Width = 2000
    Else
        grdAllLookup.Columns(j).Width = 2000
    End If
    grdWidth = grdWidth + grdAllLookup.Columns(j).Width
    
  'hide unneeded columns
  For i = j + 1 To grdAllLookup.Columns.count - 1
    grdAllLookup.Columns(i).Visible = False
  Next

    grdAllLookup.Width = grdWidth + 650
    Me.Width = grdAllLookup.Width + 100
    'MsgBox grdAllLookup.Width
    If Me.Width < 4230 Then
        Me.Width = 4230
        grdAllLookup.Width = 4130
    ElseIf Me.Width > 9230 Then
        Me.Width = 9230
        grdAllLookup.Width = 9230 - 100
    End If
     Me.Left = (Screen.Width - Me.Width) / 2
     Me.Top = (Screen.Height - Me.Height) / 2
     grdAllLookup.Left = (Me.ScaleWidth - grdAllLookup.Width) / 2
    
     Label1 = grdAllLookup.Columns(0).Caption
     AllLookup.Show vbModal

Exit Sub
NoDataToPublish:
    Exit Sub
End Sub

Private Sub Form_LostFocus()
   Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
If dbTemp Is Nothing Then
Else
    dbTemp.Close
    Set dbTemp = Nothing
End If
    Set AllLookup = Nothing
End Sub

Private Sub grdAllLookup_DblClick()
    CmdSelect_Click
End Sub

Private Sub grdAllLookup_HeadClick(ByVal ColIndex As Integer)
If MovetoCommand = True Then Exit Sub 'this is for selection task
If ADOprimaryrs.RecordCount = 0 Then Exit Sub

    Label1 = grdAllLookup.Columns(ColIndex).Caption
    WhichField = grdAllLookup.Columns(ColIndex).DataField
    ADOprimaryrs.Close
    Set ADOprimaryrs = Nothing
    Set ADOprimaryrs = New ADODB.Recordset
    ADOprimaryrs.Open TempStr & " ORDER BY [" & WhichField & "]", dbTemp, adOpenKeyset, adLockReadOnly, adCmdText
    Set grdAllLookup.DataSource = ADOprimaryrs
  
Dim TotalChar, i, j As Integer
Dim storedString As String
Dim SQLstatement As String
Dim grdDataGridHeader As String
Dim grdWidth As Integer
 
 grdDataGridHeader = GridHeader
 grdWidth = 0
  'Change The datagrid columnheader name
TotalChar = Len(grdDataGridHeader)
    For i = 1 To TotalChar

         If Mid(grdDataGridHeader, i, 2) = "//" Then
            grdDataGridHeader = Right(grdDataGridHeader, TotalChar - i - 1)
            TotalChar = Len(grdDataGridHeader)
            grdAllLookup.Columns(j).Caption = storedString
            
            If Len(Trim(ADOprimaryrs("" & grdAllLookup.Columns(j).DataField & ""))) < 8 Then
                grdAllLookup.Columns(j).Width = 800
            ElseIf Len(Trim(ADOprimaryrs("" & grdAllLookup.Columns(j).DataField & ""))) < 30 Then
                grdAllLookup.Columns(j).Width = 2000
            Else
                grdAllLookup.Columns(j).Width = 2000
            End If
            grdWidth = grdWidth + grdAllLookup.Columns(j).Width
            
            j = j + 1
            storedString = ""
            i = 1
            If TotalChar = 0 Then Exit For
         End If
         storedString = Left(grdDataGridHeader, i)
         
    Next
    grdAllLookup.Columns(j).Caption = storedString
    If Len(Trim(ADOprimaryrs("" & grdAllLookup.Columns(j).DataField & ""))) < 8 Then
        grdAllLookup.Columns(j).Width = 800
    ElseIf Len(Trim(ADOprimaryrs("" & grdAllLookup.Columns(j).DataField & ""))) < 30 Then
        grdAllLookup.Columns(j).Width = 2000
    Else
        grdAllLookup.Columns(j).Width = 2000
    End If
    grdWidth = grdWidth + grdAllLookup.Columns(j).Width
    
  'hide unneeded columns
  For i = j + 1 To grdAllLookup.Columns.count - 1
    grdAllLookup.Columns(i).Visible = False
  Next
  End Sub

Private Sub grdAllLookup_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If grdAllLookup.col = -1 Or grdAllLookup.Row = -1 Then Exit Sub
    Text1 = grdAllLookup.Columns(0).Value
End Sub


Public Sub ToWhichRecord(ADOprimaryrs As ADODB.Recordset, GridName As String, grdDataGridHeader As String)

MovetoCommand = True
'give the caption name to the datagrid
    Caption = "Lookup - " & GridName
    grdAllLookup.Caption = GridName

'add or remove column to the datagrid and name for the header
Dim TotalChar, i, j As Integer
Dim storedString As String

'Important
'A DataGrid object can contain only 32767 columns, as column indices are stored
'in integers.

'If you have previously deleted a column using the Remove method, after adding
'new columns, you may need to refresh the display with the Rebind and Refresh
'methods. This instructs the DataGrid control to rebuild its internal column
'layout matrix to correctly reflect the true status of the control.
'---remove unnecessary column
    TotalChar = Len(grdDataGridHeader)
    '    If grdAllLookup.Columns.count > 0 Then
    '        i = grdAllLookup.Columns.count - 1
    '        While i <> 0
    '            grdAllLookup.Columns.Remove (i)
    '            i = i - 1
    '        Wend
    '    End If
            
  'getting the datagrid column header
  'adoPrimaryRS.Requery
  Set grdAllLookup.DataSource = ADOprimaryrs
  
  'mbDataChanged = False
     
     grdAllLookup.ClearFields
     grdAllLookup.ReBind
     grdAllLookup.Refresh
     
     'If grdDataGridHeader = "" Then TotalChar = adoPrimaryRS.Fields.count
  'Change The datagrid columnheader name
  'Exit Sub
    For i = 1 To TotalChar
         'If grdDataGridHeader = "" Then
         '   grdAllLookup.Columns(j).Caption = "a"
         If Mid(grdDataGridHeader, i, 2) = "//" Then
            grdDataGridHeader = Right(grdDataGridHeader, TotalChar - i - 1)
            TotalChar = Len(grdDataGridHeader)
            grdAllLookup.Columns(j).Caption = storedString
            If grdAllLookup.Columns(j).Caption = "-" Then
                grdAllLookup.Columns(j).Visible = False
            Else
                grdAllLookup.Columns(j).Visible = True
            End If
            j = j + 1
            storedString = ""
            i = 1
            If TotalChar = 0 Then Exit For
         End If
         storedString = Left(grdDataGridHeader, i)
  
    Next
    grdAllLookup.Columns(j).Caption = storedString
  
  'hide unneeded columns
  For i = j + 1 To grdAllLookup.Columns.count - 1
    grdAllLookup.Columns(i).Visible = False
  Next
    
    AllLookup.Show vbModal

End Sub

