VERSION 5.00
Begin VB.Form frmInstall 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Orb Installation"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5550
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   5550
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNever 
      Caption         =   "Never"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Install"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   4920
      Top             =   240
   End
   Begin VB.CommandButton cmdAlternate 
      Caption         =   "Later"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmInstall.frx":0000
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5235
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmInstall.frx":00D7
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   5235
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "To uninstall the start toolbar at anytime, right click the Windows taskbar and select ""Toolbars"" > ""Start""."
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   915
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   5235
   End
End
Attribute VB_Name = "frmInstall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'    Component  : frmInstall
'    Project    : ViOrb5
'
'    Description: On some systems (like Windows 8) there's no start button
'                 the purpose of this form then is to create a toolbar on the
'                 taskbar (a placeholder) without the need to hook the taskbar
'                 We then use that placeholder and align our start button to it
'
'
'    Modified   :
'--------------------------------------------------------------------------------
Option Explicit


'(HEX) Registry to create a blank dummy toolbar that links to nowhere on the taskbar
Private Const START_TOOLBAR As String = "0c,00,00,00,08,00,00,00,02,00,00,00,00,00,00,00,b0,e2,2b,d8,64,57,d0,11,a9,6e,00,c0,4f,d7,05,a2,22,00,1c,00,08,10,00,00,01,00,00,00,01,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,4c,00,00,00,01,14,02,00,00,00,00,00,c0,00,00,00,00,00,00,46,81,01,00,00,10,00,00,00,d2,ca,6d,af,43,54,cd,01,d2,ca,6d,af,43,54,cd,01,d2,ca,6d,af,43,54,cd,01,00,00,00,00,00,00,00,00,01,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,7b,00,14,00,1f,50,e0,4f,d0,20,ea,3a,69,10,a2,d8,08,00,2b,30,30,9d,19,00,2f,43,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,4c,00,31,00,00,00,00,00,db,40,68,48,10,00,53,74,61,72,74,00,38,00,08,00,04,00,ef,be,db,40,68,48,db,40,68,48,2a,00,00,00,0b" & _
   ",49,07,00,00,00,31,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,53,00,74,00,61,00,72,00,74,00,00,00,14,00,00,00,60,00,00,00,03,00,00,a0,58,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,da,c9,2b,f9,1d,3d,13,42,a7,8d,fc,24,b8,bf,90,51,6b,7e,87,71,2d,bf,e1,11,98,bc,bc,ae,c5,65,4f,c5,da,c9,2b,f9,1d,3d,13,42,a7,8d,fc,24,b8,bf,90,51,6b,7e,87,71,2d,bf,e1,11,98,bc,bc,ae,c5,65,4f,c5,00,00,00,00,2c,00,00,00,40,03,00,00,00,00,00,00,1e,00,00,00,00,00,00,00,00,00,00,00,28,00,00,00,00,00,00,00,00,00,00,00,01,00,00,00,aa,4f,28,68,48,6a,d0,11,8c,78,00,c0,4f,d9,18,b4,ee,04,00,00,40,0d,00,00,00,00,00,00,28,00,00,00,00,00,00,00,00,00,00,00,28,00,00,00,00,00,00,00,01,00,00,00"
                                        
Private Declare Function RegCreateKey _
                Lib "advapi32.dll" _
                Alias "RegCreateKeyA" (ByVal hKey As Long, _
                                       ByVal lpSubKey As String, _
                                       phkResult As Long) As Long

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Private Declare Function RegSetValueEx _
                Lib "advapi32.dll" _
                Alias "RegSetValueExA" (ByVal hKey As Long, _
                                        ByVal lpValueName As String, _
                                        ByVal Reserved As Long, _
                                        ByVal dwType As Long, _
                                        lpData As Any, _
                                        ByVal cbData As Long) As Long

Private Const REG_BINARY = 3

Private Function HexToBin(theHex As String) As Byte()

    Dim hexArr()     As String

    Dim hexIndex     As Long

    Dim returnByte() As Byte

    hexArr = Split(theHex, ",")
    ReDim returnByte(UBound(hexArr))

    For hexIndex = LBound(hexArr) To UBound(hexArr)
        returnByte(hexIndex) = Val("&H" & hexArr(hexIndex) & "&")
    Next
    
    HexToBin = returnByte
End Function

Private Function WriteBinaryToRegistry(hKey As Long, _
                                       strPath As String, _
                                       strValue As String, _
                                       binData() As Byte) As Boolean

    On Error GoTo ErrorHandler

    Dim keyhand As Long

    Dim r       As Long

    r = RegCreateKey(hKey, strPath, keyhand)

    If r = 0 Then
        r = RegSetValueEx(keyhand, strValue, 0, REG_BINARY, binData(0), UBound(binData) + 1)
        r = RegCloseKey(keyhand)
    End If
    
    WriteBinaryToRegistry = (r = 0)

    Exit Function

ErrorHandler:
    WriteBinaryToRegistry = False

    Exit Function
    
End Function

Private Sub cmdAlternate_Click()
    Unload Me
End Sub

Private Sub cmdNever_Click()
    WriteRegistryString AppSettingsRegistryPath & "noorb_warning", "1"
    Unload Me
End Sub

Private Sub cmdNext_Click()
    cmdNext.Enabled = False

    WriteBinaryToRegistry HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Streams\Desktop\", "TaskbarWinXP", HexToBin(START_TOOLBAR)
    WriteRegistryInteger AppSettingsRegistryPath & Default_Orb_Name & "\nudgeX", -8
    WriteRegistryInteger AppSettingsRegistryPath & Default_Orb_Name & "\nudgeY", -2
    
    KillProcess "explorer"
    Sleep 1000
    
    If (ProcessCount("explorer") = 0) Then
        ShellCommand "explorer.exe"
    End If
    
    cmdNext.Enabled = True
End Sub

Private Sub Form_Load()

    If g_Windows8 And Not g_Windows81 Then
        cmdNever.Enabled = False
    End If

End Sub

Private Sub Timer1_Timer()
    TaskbarHelper.UpdatehWnds
    
    If IsWindow(TaskbarHelper.g_viOrbToolbar) = APITRUE Then
        Unload Me
    End If

End Sub


