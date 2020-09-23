VERSION 5.00
Begin VB.Form Main_About 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About NoSecret Accounting Tech."
   ClientHeight    =   4050
   ClientLeft      =   5970
   ClientTop       =   3015
   ClientWidth     =   3735
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   3735
   Tag             =   "1068"
   WhatsThisHelp   =   -1  'True
   Begin VB.Timer Timer1 
      Interval        =   30
      Left            =   2880
      Top             =   3120
   End
   Begin VB.PictureBox picMajor 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   2655
      Left            =   0
      ScaleHeight     =   2595
      ScaleWidth      =   3555
      TabIndex        =   2
      Top             =   0
      Width           =   3615
      Begin VB.PictureBox picMove 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   11520
         Index           =   0
         Left            =   120
         ScaleHeight     =   11520
         ScaleWidth      =   3495
         TabIndex        =   3
         Top             =   1560
         Width           =   3495
         Begin VB.PictureBox Picture2 
            Height          =   1935
            Left            =   0
            Picture         =   "Main_About.frx":0000
            ScaleHeight     =   1875
            ScaleWidth      =   915
            TabIndex        =   16
            Top             =   4080
            Width           =   975
         End
         Begin VB.PictureBox Picture1 
            Height          =   1815
            Left            =   2520
            Picture         =   "Main_About.frx":0D78
            ScaleHeight     =   1755
            ScaleWidth      =   795
            TabIndex        =   15
            Top             =   2400
            Width           =   855
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   $"Main_About.frx":25DC
            ForeColor       =   &H0000FFFF&
            Height          =   1815
            Left            =   0
            TabIndex        =   13
            Top             =   9720
            Width           =   3135
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Microsoft Corp. for distribution of TransTBWrapper which is used by us in the toolbar."
            ForeColor       =   &H0000FFFF&
            Height          =   615
            Left            =   0
            TabIndex        =   12
            Top             =   9000
            Width           =   3015
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Special Thanks to:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF00&
            Height          =   255
            Left            =   0
            TabIndex        =   11
            Top             =   8520
            Width           =   3015
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   $"Main_About.frx":277A
            ForeColor       =   &H0000FFFF&
            Height          =   1095
            Left            =   0
            TabIndex        =   10
            Top             =   6240
            Width           =   3375
         End
         Begin VB.Image Image5 
            BorderStyle     =   1  'Fixed Single
            Height          =   765
            Left            =   720
            Picture         =   "Main_About.frx":287C
            Top             =   7560
            Width           =   2010
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   $"Main_About.frx":2AB3
            ForeColor       =   &H0000FFFF&
            Height          =   1095
            Left            =   1080
            TabIndex        =   9
            Top             =   4560
            Width           =   2295
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   $"Main_About.frx":2B65
            ForeColor       =   &H0000FFFF&
            Height          =   1095
            Left            =   0
            TabIndex        =   8
            Top             =   2880
            Width           =   2295
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "NoSecret Accounting"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Index           =   1
            Left            =   0
            TabIndex        =   7
            Top             =   0
            Width           =   3255
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "(C)opyright 2000, NoSecret Accounting Tech. All Rights Reserved"
            ForeColor       =   &H0000FFFF&
            Height          =   495
            Index           =   2
            Left            =   0
            TabIndex        =   6
            Top             =   1680
            Width           =   3255
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   $"Main_About.frx":2C0E
            ForeColor       =   &H0000FFFF&
            Height          =   975
            Index           =   4
            Left            =   0
            TabIndex        =   5
            Top             =   480
            Width           =   2535
         End
         Begin VB.Image Image2 
            Height          =   630
            Left            =   2640
            Picture         =   "Main_About.frx":2CAE
            Top             =   720
            Width           =   660
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Trademark and Copyright"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF00&
            Height          =   255
            Left            =   0
            TabIndex        =   4
            Top             =   2400
            Width           =   2895
         End
      End
      Begin VB.PictureBox picMove 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   10815
         Index           =   1
         Left            =   120
         ScaleHeight     =   10815
         ScaleWidth      =   3495
         TabIndex        =   14
         Top             =   -8160
         Visible         =   0   'False
         Width           =   3495
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Gone but not forgotten:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF80&
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   41
            Top             =   6240
            Width           =   2415
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Technical Support"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF80&
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   40
            Top             =   5160
            Width           =   2415
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Developer Support"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF80&
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   39
            Top             =   4920
            Width           =   2415
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Quality Engineering"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF80&
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   38
            Top             =   4680
            Width           =   2415
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "User Interface"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF80&
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   37
            Top             =   3840
            Width           =   2415
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Database Design"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF80&
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   36
            Top             =   3000
            Width           =   2415
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Database Flow Design"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF80&
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   35
            Top             =   2760
            Width           =   2415
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Application Flow Design"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF80&
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   34
            Top             =   1920
            Width           =   2415
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Engineering"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF80&
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   33
            Top             =   1680
            Width           =   2415
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Beta Coordination"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF80&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   32
            Top             =   1440
            Width           =   2415
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Management"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF80&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   31
            Top             =   120
            Width           =   2415
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "MAJOR THANKS TO ALL"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   120
            TabIndex        =   30
            Top             =   9000
            Width           =   3255
         End
         Begin VB.Label Label21 
            BackColor       =   &H0000FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Our Sincere thanks and apologies to anyone who deserves credit but fail to appear in the list."
            ForeColor       =   &H0000FFFF&
            Height          =   615
            Left            =   120
            TabIndex        =   29
            Top             =   7920
            Width           =   3135
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "3. Harun Saad"
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   240
            TabIndex        =   28
            Top             =   7080
            Width           =   2295
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "4. Abdul Hajat Ishak"
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   240
            TabIndex        =   27
            Top             =   7320
            Width           =   2535
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "2. Nor Azli Marsudin"
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   240
            TabIndex        =   26
            Top             =   6840
            Width           =   2415
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "1. Shanizam Akma Md. Shabudin"
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   240
            TabIndex        =   25
            Top             =   6600
            Width           =   2535
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "2. Mohd Yusri Shaharin"
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   240
            TabIndex        =   24
            Top             =   5760
            Width           =   1935
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "1. Mohd Razi Abdul Latif"
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   240
            TabIndex        =   23
            Top             =   5520
            Width           =   2175
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "1. Mohd Yusri Shaharin"
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   240
            TabIndex        =   22
            Top             =   3360
            Width           =   2295
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "1.Mohd Razi Abdul Latif"
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   240
            TabIndex        =   21
            Top             =   4200
            Width           =   2175
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "1. Mohd Razi Abdul Latif"
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   240
            TabIndex        =   20
            Top             =   2280
            Width           =   1935
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "3. Nushirwan Abdul Rahim"
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   240
            TabIndex        =   19
            Top             =   960
            Width           =   2055
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "2. Suhaila Yacob"
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Top             =   720
            Width           =   2295
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "1. Mohd Fikri"
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   240
            TabIndex        =   17
            Top             =   480
            Width           =   2535
         End
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   240
      TabIndex        =   0
      Tag             =   "1070"
      Top             =   3600
      Width           =   1467
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&System Info..."
      Height          =   345
      Left            =   2040
      TabIndex        =   1
      Tag             =   "1069"
      Top             =   3600
      Width           =   1452
   End
   Begin VB.Image Image3 
      Height          =   825
      Left            =   1320
      Picture         =   "Main_About.frx":3121
      Top             =   2760
      Width           =   825
   End
End
Attribute VB_Name = "Main_About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Reg Key Security Options...
Const KEY_ALL_ACCESS = &H2003F
                                          

' Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number


Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"


Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Dim i1 As Integer

Private Sub Form_Load()
    picMajor.Left = (Me.ScaleWidth - picMajor.Width) / 2
    picMove(0).Top = 120
    picMove(0).Left = 120
    picMove(1).Top = 120
    picMove(1).Left = 120
    i1 = 0
End Sub

Private Sub cmdSysInfo_Click()
        Call StartSysInfo
End Sub


Private Sub cmdOK_Click()
        Unload Me
End Sub


Public Sub StartSysInfo()
    'On Error GoTo SysInfoErr


        Dim rc As Long
        Dim SysInfoPath As String
        

        ' Try To Get System Info Program Path\Name From Registry...
        If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
        ' Try To Get System Info Program Path Only From Registry...
        ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
                ' Validate Existance Of Known 32 Bit File Version
                If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
                        SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
                        

                ' Error - File Can Not Be Found...
                Else
                        GoTo SysInfoErr
                End If
        ' Error - Registry Entry Can Not Be Found...
        Else
                GoTo SysInfoErr
        End If
        

        Call Shell(SysInfoPath, vbNormalFocus)
        

        Exit Sub
SysInfoErr:
        MsgBox "System Information Is Unavailable At This Time", vbOKOnly
End Sub


Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
        Dim i As Long                                           ' Loop Counter
        Dim rc As Long                                          ' Return Code
        Dim hKey As Long                                        ' Handle To An Open Registry Key
        Dim hDepth As Long                                      '
        Dim KeyValType As Long                                  ' Data Type Of A Registry Key
        Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
        Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
        '------------------------------------------------------------
        ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
        '------------------------------------------------------------
        rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
        

        If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
        

        tmpVal = String$(1024, 0)                             ' Allocate Variable Space
        KeyValSize = 1024                                       ' Mark Variable Size
        

        '------------------------------------------------------------
        ' Retrieve Registry Key Value...
        '------------------------------------------------------------
        rc = RegQueryValueEx(hKey, SubKeyRef, 0, KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                                                

        If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
        

        tmpVal = VBA.Left(tmpVal, InStr(tmpVal, VBA.Chr(0)) - 1)
        '------------------------------------------------------------
        ' Determine Key Value Type For Conversion...
        '------------------------------------------------------------
        Select Case KeyValType                                  ' Search Data Types...
        Case REG_SZ                                             ' String Registry Key Data Type
                KeyVal = tmpVal                                     ' Copy String Value
        Case REG_DWORD                                          ' Double Word Registry Key Data Type
                For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
                        KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
                Next
                KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
        End Select
        

        GetKeyValue = True                                      ' Return Success
        rc = RegCloseKey(hKey)                                  ' Close Registry Key
        Exit Function                                           ' Exit
        

GetKeyError:    ' Cleanup After An Error Has Occured...
        KeyVal = ""                                             ' Set Return Val To Empty String
        GetKeyValue = False                                     ' Return Failure
        rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function

Private Sub Form_Unload(Cancel As Integer)
   Timer1.Enabled = False
   Set Main_About = Nothing
End Sub

Private Sub Timer1_Timer()
Dim Movetop As Integer

    i1 = i1 + 1
If i1 > 100 Then
    i1 = 100
    Movetop = 10
      If picMove(0).Top > -9240 And picMove(1).Visible = False Then
        picMove(0).Top = picMove(0).Top - Movetop
        picMove(1).Top = 2400
      ElseIf picMove(0).Top <= -9240 And picMove(0).Top > -11640 Then
        picMove(1).Visible = True
        picMove(1).Top = picMove(1).Top - Movetop
        picMove(0).Top = picMove(0).Top - Movetop
      ElseIf picMove(0).Top = -11640 And picMove(0).Visible = True Then
        picMove(0).Visible = False
      ElseIf picMove(1).Top > -8400 And picMove(0).Visible = False Then
        picMove(1).Top = picMove(1).Top - Movetop
        picMove(0).Top = 2400
      ElseIf picMove(1).Top <= -8400 And picMove(1).Top > -10800 Then
        picMove(0).Visible = True
        picMove(0).Top = picMove(0).Top - Movetop
        picMove(1).Top = picMove(1).Top - Movetop
      ElseIf picMove(1).Top = -10800 And picMove(1).Visible = True Then
        picMove(1).Visible = False
      End If
End If
End Sub
