VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Menu_Calendar 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Appointment"
   ClientHeight    =   4200
   ClientLeft      =   2700
   ClientTop       =   6330
   ClientWidth     =   4740
   FillColor       =   &H0000C000&
   ForeColor       =   &H00000000&
   Icon            =   "Menu_Calendar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "Menu_Calendar.frx":030A
   ScaleHeight     =   4200
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   4095
      Left            =   0
      ScaleHeight     =   4095
      ScaleWidth      =   4695
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.CommandButton cmdAppointment 
         Height          =   555
         Left            =   4080
         Picture         =   "Menu_Calendar.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Appointment"
         Top             =   120
         Width           =   615
      End
      Begin VB.TextBox txtYear 
         Height          =   320
         Left            =   2160
         TabIndex        =   4
         Top             =   120
         Width           =   1215
      End
      Begin VB.ComboBox cbMonth 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   120
         Width           =   2055
      End
      Begin VB.PictureBox picMonth 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H0000FF00&
         Height          =   3255
         Left            =   120
         ScaleHeight     =   3255
         ScaleWidth      =   4455
         TabIndex        =   1
         Top             =   960
         Width           =   4455
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   2
         Top             =   600
         Width           =   4695
      End
   End
   Begin MSDataGridLib.DataGrid grddatagrid 
      Height          =   3615
      Left            =   4800
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   6376
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
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
      Caption         =   "Appointment"
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "DIARY Owner"
         Caption         =   "DIARY Owner"
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
         DataField       =   "DIARY Date"
         Caption         =   "DIARY Date"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "MM/dd/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "DIARY Time"
         Caption         =   "Time"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "h:nn AM/PM"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "DIARY Description"
         Caption         =   "Description"
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
         BeginProperty Column00 
            DividerStyle    =   1
            Object.Visible         =   0   'False
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column01 
            Object.Visible         =   0   'False
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column03 
            WrapText        =   -1  'True
            ColumnWidth     =   4050.142
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Appointment"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   0
      Left            =   4800
      TabIndex        =   7
      Top             =   120
      Width           =   5775
   End
End
Attribute VB_Name = "Menu_Calendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ADOprimaryrs As ADODB.Recordset
Dim db As ADODB.Connection

'Grid dimensions for days
Private Const GRID_ROWS = 7
Private Const GRID_COLS = 7
'Private Const MEM_Count = 40

'Private variables
Private m_CurrDate As Date ', m_bAcceptChange As Boolean
Private m_nGridWidth As Integer, m_nGridHeight As Integer
Private NewLoad As Boolean
Dim OnCall As Boolean
Dim TargetTxt As Integer
Private grdOnAddNew As Boolean

'Public function: If user selects date, sets UserDate to selected
'date and returns True. Otherwise, returns False.
Private Function GetDate(Userdate As Date, Optional title) As Boolean

    'Store user-specified date
    m_CurrDate = Userdate
    
    'Use caller-specified caption if any
    If Not IsMissing(title) Then
        Caption = title
    End If

    'Return selected date
    'If m_bAcceptChange Then
    '    Userdate = m_CurrDate
    'End If

    'Return value indicates if date was selected
    'GetDate = m_bAcceptChange
End Function

Private Sub cbMonth_Click()
If NewLoad Then Exit Sub
Dim i As Integer
    For i = 0 To 11
        If cbMonth.Text = cbMonth.List(i) Then Exit For
    Next
    SetNewDate Format(Format(Str(i + 1) & "/1/" & txtYear, "mm/dd/yyyy"), "mm/dd/yyyy")
    
    If grdDataGrid.Visible = True Then OpenDB "SELECT * FROM [Diary] WHERE [DIARY Owner]='" & AppLoginName & "' AND [DIARY Date]=#" & FormatDate(m_CurrDate) & "#"
End Sub

Private Sub cmdAppointment_Click()
    OpenDB "SELECT * FROM [Diary] WHERE [DIARY Owner]='" & AppLoginName & "' AND [DIARY Date]=#" & FormatDate(m_CurrDate) & "#"
    grdDataGrid.Visible = True
    Me.Width = 10710
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim NewDate As Date
    Dim MoveDate As Boolean
    'Board.Visible = True
    Select Case KeyCode
        Case vbKeyRight
            NewDate = DateAdd("d", 1, m_CurrDate)
            MoveDate = True
        Case vbKeyLeft
            NewDate = DateAdd("d", -1, m_CurrDate)
            MoveDate = True
        Case vbKeyDown
            NewDate = DateAdd("ww", 1, m_CurrDate)
            MoveDate = True
        Case vbKeyUp
            NewDate = DateAdd("ww", -1, m_CurrDate)
            MoveDate = True
        Case vbKeyPageDown
            NewDate = DateAdd("m", 1, m_CurrDate)
            MoveDate = True
        Case vbKeyPageUp
            NewDate = DateAdd("m", -1, m_CurrDate)
            MoveDate = True
        'Case vbKeyReturn
            'm_bAcceptChange = True
        '    Unload Me
        '    Exit Sub
        Case vbKeyEscape
            Unload Me
            Exit Sub
        'Case Else
        '    Exit Sub
    End Select
If MoveDate Then
    SetNewDate NewDate
    If grdDataGrid.Visible = True Then OpenDB "SELECT * FROM [Diary] WHERE [DIARY Owner]='" & AppLoginName & "' AND [DIARY Date]=#" & FormatDate(m_CurrDate) & "#"
    'KeyCode = 0
End If
End Sub

'Form initialization
Private Sub Form_Load()
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open gblBasicADOProvider

NewLoad = True
Dim i As Integer
    'Center form on screen
    'Width = 2400
    'Height = 2190
    'Board.Left = 0: Board.top = 0
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    'show current date
    GetDate Date
    For i = 1 To 12
        cbMonth.List(i - 1) = Format(Str(i) & "/1/" & Format(Date, "yyyy"), "mmmm")
    Next
    cbMonth.Text = Format(Date, "mmmm")
    'Calculate calendar grid measurements
    m_nGridWidth = ((picMonth.ScaleWidth - Screen.TwipsPerPixelX) \ GRID_COLS)
    m_nGridHeight = ((picMonth.ScaleHeight - Screen.TwipsPerPixelY) \ GRID_ROWS)
    'show the current date on screen
    txtYear.Text = Format(m_CurrDate, "yyyy")
    If grdDataGrid.Visible = True Then OpenDB "SELECT * FROM [Diary] WHERE [DIARY Owner]='" & AppLoginName & "' AND [DIARY Date]=#" & FormatDate(m_CurrDate) & "#"
    NewLoad = False
    'm_bAcceptChange = False
End Sub

Private Sub OpenDB(SQLstatement As String)
ShowStatus True
 Set grdDataGrid.DataSource = Nothing
 If ADOprimaryrs Is Nothing Then
 Else
    If ADOprimaryrs.RecordCount > 0 Then ADOprimaryrs.Update
    ADOprimaryrs.Close
    Set ADOprimaryrs = Nothing
 End If
 
 Set ADOprimaryrs = New ADODB.Recordset
 ADOprimaryrs.Open SQLstatement, db, adOpenKeyset, adLockOptimistic, adCmdText
 Set grdDataGrid.DataSource = ADOprimaryrs
 'MsgBox SQLstatement & "   " & ADOprimaryRS.RecordCount
ShowStatus False
grdOnAddNew = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
If ADOprimaryrs Is Nothing Then
Else
    If ADOprimaryrs.RecordCount > 0 Then ADOprimaryrs.Update
    ADOprimaryrs.Close
    Set ADOprimaryrs = Nothing
End If
    db.Close
    Set db = Nothing

    Set Menu_Calendar = Nothing
End Sub

Private Sub grdDataGrid_AfterColEdit(ByVal ColIndex As Integer)
      SendKeys ("{ENTER}")
      SendKeys ("{Left}")
      SendKeys ("{Right}")
End Sub

Private Sub grdDataGrid_Error(ByVal DataError As Integer, Response As Integer)
    Select Case DataError
    Case 13
        MsgBox "Please insert only time Value HH:MM."
        Response = 0
    End Select
End Sub

Private Sub grdDataGrid_OnAddNew()
    grdDataGrid.Columns(0) = AppLoginName
    grdDataGrid.Columns(1) = FormatDate(m_CurrDate)
    grdOnAddNew = True
End Sub

Private Sub grdDataGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If grdOnAddNew = True Then
    'If ADOprimaryrs.RecordCount = 0 Then grddatagrid.Row = grddatagrid.Row + 1
    grdOnAddNew = False
    ADOprimaryrs.Update
    OpenDB "SELECT * FROM [Diary] WHERE [DIARY Owner]='" & AppLoginName & "' AND [DIARY Date]=#" & FormatDate(m_CurrDate) & "#"
    If ADOprimaryrs.RecordCount > 1 Then ADOprimaryrs.MoveLast
End If
End Sub

'Double-click accepts current date
Private Sub picMonth_DblClick()
If OnCall = False Then
    Unload Me
    'Exit Sub
End If
Select Case TargetTxt
Case 1000   ' calling from frm_AR_Quote_Entry for txtfields(20)
   fMainForm.ActiveForm.txtfields(20) = FormatDate(m_CurrDate)
Case 1001   ' calling from frm_AR_Order_Entry for txtfields(6)
   fMainForm.ActiveForm.txtfields(6) = FormatDate(m_CurrDate)
Case 1002   ' calling from frm_AR_Order_Entry for txtfields(7)
   fMainForm.ActiveForm.txtfields(7) = FormatDate(m_CurrDate)
   fMainForm.ActiveForm.SetDueDate
Case 1003   ' calling from frm_SYS_Setup_Period for txtfields(26)
   fMainForm.ActiveForm.txtPeriods(26) = FormatDate(m_CurrDate)
Case 1004   ' calling from frm_AP_Purchase_Entry for txtfields(27)
   fMainForm.ActiveForm.txtPeriods(27) = FormatDate(m_CurrDate)
Case 1010   ' calling from frm_AP_Purchase_Entry for txtfields(7)
   fMainForm.ActiveForm.txtfields(7) = FormatDate(m_CurrDate)
Case 1011   ' calling from frm_AP_Purchase_Entry for txtfields(4)
   fMainForm.ActiveForm.txtfields(4) = FormatDate(m_CurrDate)
Case 1012   ' calling from frm_AP_Purchase_Entry for txtfields(14)
   fMainForm.ActiveForm.txtfields(14) = FormatDate(m_CurrDate)
Case 1020   ' calling from frm_AP_Receiving_Entry for grdDataGrid.Columns(6)
   fMainForm.ActiveForm.grdDataGrid.Columns(6) = FormatDate(m_CurrDate)
Case 1030   ' calling from frm_AP_Receiving_Entry for txtFields(31)
   fMainForm.ActiveForm.txtfields(31) = FormatDate(m_CurrDate)
Case 1302
   fMainForm.ActiveForm.txtfields(0) = FormatDate(m_CurrDate)
Case 1322   'calling from frm_Bank_Transaction for txtfields(3)
   fMainForm.ActiveForm.txtfields(3) = FormatDate(m_CurrDate)
Case 1422   'calling from frm_Bank_Reconciliation for txtfields(5)
   fMainForm.ActiveForm.txtfields(5) = FormatDate(m_CurrDate)
Case 1432   'calling from frm_Bank_Reconciliation for txtfields(4)
   fMainForm.ActiveForm.txtfields(4) = FormatDate(m_CurrDate)
Case 1500   'calling from frm_Pay_Employees for txtFieldsDate(0)
   fMainForm.ActiveForm.txtFieldsDate(0) = FormatDate(m_CurrDate)
Case 1510   'calling from frm_Pay_Employees for txtFieldsDate(1)
   fMainForm.ActiveForm.txtFieldsDate(1) = FormatDate(m_CurrDate)
Case 1520   'calling from frm_Pay_Employees for txtFieldsDate(2)
   fMainForm.ActiveForm.txtFieldsDate(2) = FormatDate(m_CurrDate)
Case 1525   'calling from frm_Pay_Employees for grddatagrid.columns(6)
   fMainForm.ActiveForm.grdDataGrid.Columns(6) = FormatDate(m_CurrDate)
Case 1535   'calling from frm_Pay_Employees for grddatagrid.columns(6)
   fMainForm.ActiveForm.txtfields(4) = FormatDate(m_CurrDate)
Case 1545   'calling from frm_Pay_Employees for txtFields(1)
   fMainForm.ActiveForm.txtCommission(1) = FormatDate(m_CurrDate)
Case 1555   'calling from frm_Pay_Employees for txtFields(10)
   fMainForm.ActiveForm.txtfields(10) = FormatDate(m_CurrDate)
Case 1556   'calling from frm_Pay_Employees for txtFields(11)
   fMainForm.ActiveForm.txtfields(11) = FormatDate(m_CurrDate)
Case 1565   'calling from frm_LIST_Credit_Cards for txtFields(3)
   fMainForm.ActiveForm.txtfields(3) = FormatDate(m_CurrDate)
Case 1600   'calling from frm_LIST_Credit_Cards for txtFields(3)
   fMainForm.ActiveForm.txtfields(4) = FormatDate(m_CurrDate)
Case 1610   'calling from frm_LIST_Credit_Cards for txtFields(3)
   fMainForm.ActiveForm.txtfields(12) = FormatDate(m_CurrDate)
Case 1620   'calling from frm_SYS_Setup_Employee for txt(1)
   fMainForm.ActiveForm.txt(1) = FormatDate(m_CurrDate)
Case 1630   'calling from frm_SYS_Setup_Employee for txt(2)
   fMainForm.ActiveForm.txt(2) = FormatDate(m_CurrDate)
Case 1640   'calling from frm_Check_Management for txtfields(1)
   fMainForm.ActiveForm.txtfields(1) = FormatDate(m_CurrDate)
Case 1645   'calling from frm_AR_Customer for txtfields(17)
   fMainForm.ActiveForm.txtfields(17) = FormatDate(m_CurrDate)
Case 1650   'calling from frm_AR_Customer for txtfields(16)
   fMainForm.ActiveForm.txtfields(16) = FormatDate(m_CurrDate)
Case 1660   'calling from frm_Recurring for grdDataGrid2.columns(1)
   fMainForm.ActiveForm.grdDataGrid2.Columns(1) = FormatDate(m_CurrDate)
Case 1670   'calling from frm_Recurring for grdDataGrid2.columns(1)
   fMainForm.ActiveForm.txtfields(32) = FormatDate(m_CurrDate)
Case 1680   'calling from frm_Recurring for grdDataGrid2.columns(1)
   fMainForm.ActiveForm.txtfields(33) = FormatDate(m_CurrDate)
End Select
Unload Me
End Sub

' Select the date by mouse
Private Sub picMonth_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer, MaxDay As Integer
    
    'Determine which date is being clicked
    i = Weekday(DateSerial(Year(m_CurrDate), Month(m_CurrDate), 1)) - 1
    i = (((X \ m_nGridWidth) + 1) + ((Y \ m_nGridHeight) * GRID_COLS)) - i
    'Get last day of current month
    MaxDay = Day(DateAdd("d", -1, DateSerial(Year(m_CurrDate), Month(m_CurrDate) + 1, 1)))
    
    If i >= 1 And i <= MaxDay Then
        SetNewDate DateSerial(Year(m_CurrDate), Month(m_CurrDate), i)
    End If
    If grdDataGrid.Visible = True Then OpenDB "SELECT * FROM [Diary] WHERE [DIARY Owner]='" & AppLoginName & "' AND [DIARY Date]=#" & FormatDate(m_CurrDate) & "#"
End Sub

'Changes the selected date
Private Sub SetNewDate(NewDate As Date)
    If Month(m_CurrDate) = Month(NewDate) And Year(m_CurrDate) = Year(NewDate) Then
        DrawSelectionBox False
'******************************************************
'Setting a new date for input --- m_CurrDate
'******************************************************
        m_CurrDate = NewDate
        DrawSelectionBox True
    Else
        m_CurrDate = NewDate
        picMonth_Paint
    End If
End Sub

'Here's the calendar paint handler; displayes the calendar days
Private Sub picMonth_Paint()
    Dim i As Integer, j As Integer, X As Integer, Y As Integer
    Dim NumDays As Integer, CurrPos As Integer, bCurrMonth As Boolean
    Dim MonthStart As Date, Buffer As String
    
    'Determine if this month is today's month
    If Month(m_CurrDate) = Month(Date) And Year(m_CurrDate) = Year(Date) Then
        bCurrMonth = True
    End If

    'Get first date in the month
    MonthStart = DateSerial(Year(m_CurrDate), Month(m_CurrDate), 1)
    
    'Number of days in the month
    NumDays = DateDiff("d", MonthStart, DateAdd("m", 1, MonthStart))

    'Get first weekday in the month (0 - based)
    j = Weekday(MonthStart) - 1
    
    'Tweak for 1-based For/Next index
    j = j - 1

    'Show current month/year
    'lblMonth = Format$(m_CurrDate, "mmmm yyyy")
    
    'Clear existing data
    picMonth.Cls

    'Display dates for current month
    For i = 1 To NumDays
        CurrPos = i + j
        X = (CurrPos Mod GRID_COLS) * m_nGridWidth
        Y = (CurrPos \ GRID_COLS) * m_nGridHeight
        'Show date as bold if today's date
        If bCurrMonth And i = Day(Date) Then
            picMonth.Font.Bold = True
            picMonth.ForeColor = &HFFFF&
        Else
            picMonth.Font.Bold = False
            picMonth.ForeColor = vbBlack
        End If
        'Center date within "date cell"
        Buffer = CStr(i)
        picMonth.CurrentX = X + ((m_nGridWidth - picMonth.TextWidth(Buffer)) / 2)
        picMonth.CurrentY = Y + ((m_nGridHeight - picMonth.TextHeight(Buffer)) / 2)
        'Print date
        picMonth.Print Buffer;
    Next i
    'Indicate selected date
    DrawSelectionBox True
End Sub

'Draw or clears the selection box around the current date
Private Sub DrawSelectionBox(bSelected As Boolean)
    Dim clrTopLeft As Long, clrBottomRight As Long
    Dim i As Integer, X As Integer, Y As Integer

    'Set highlight and shadow colors
    If bSelected Then
        clrTopLeft = &HC0FFC0
        clrBottomRight = &H8000&
    Else
        clrTopLeft = picMonth.BackColor
        clrBottomRight = picMonth.BackColor
    End If
    
    'Compute location for current date
    lbl = Format(m_CurrDate, "dddd") & "    " & Format(m_CurrDate, "mm/dd/yyyy")
    i = Weekday(DateSerial(Year(m_CurrDate), Month(m_CurrDate), 1)) - 1
    i = i + (Day(m_CurrDate) - 1)
    X = (i Mod GRID_COLS) * m_nGridWidth
    Y = (i \ GRID_COLS) * m_nGridHeight
    'FindDay X
    'Draw box around date
    picMonth.Line (X, Y + m_nGridHeight)-Step(0, -m_nGridHeight), clrTopLeft
    picMonth.Line -Step(m_nGridWidth, 0), clrTopLeft
    picMonth.Line -Step(0, m_nGridHeight), clrBottomRight '= white
    picMonth.Line -Step(-m_nGridWidth, 0), clrBottomRight '= white
End Sub

Private Sub txtYear_LostFocus()
    If IsNumeric(txtYear) = False Then
        txtYear.Text = Format(m_CurrDate, "yyyy")
    Else
        txtYear.Text = Int(txtYear)
    End If
    cbMonth_Click
End Sub

Public Sub WhoCallMe(GotCall As Boolean, CallText As Integer)
    OnCall = GotCall
    TargetTxt = CallText
    cmdAppointment.Enabled = False
    Menu_Calendar.Show vbModal
End Sub

