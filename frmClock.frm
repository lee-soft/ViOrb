VERSION 5.00
Begin VB.Form frmClock 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "00:00 AM"
   ClientHeight    =   210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   210
   Icon            =   "frmClock.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   210
   ScaleWidth      =   210
   ShowInTaskbar   =   0   'False
   Begin VB.Timer timExit 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   720
      Top             =   840
   End
End
Attribute VB_Name = "frmClock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'    Component  : frmClock
'    Project    : ViOrb5
'
'    Description: A form for the system tray icon
'                 TODO: Change this form name to something more meaningful
'
'    Modified   :
'--------------------------------------------------------------------------------
Option Explicit

Private m_PopupSystemMenu As clsMenu

Private nid               As NOTIFYICONDATA

Public Sub PopupSystemMenu()

    Select Case m_PopupSystemMenu.ShowMenu(Me.hWnd)
      
        Case 1
            timExit.Enabled = True
    
        Case 2
            frmSplash.Show

            Exit Sub
        
        Case 3
            frmOptions.Show

            Exit Sub
        
        Case 4
            RePositionOrb
        
        Case 5
            ResetPosition
        
        Case 6
            frmOptions.cmdBrowse_Click
      
    End Select

End Sub

Sub RePositionOrb()

    If IsWindow(TaskbarHelper.g_viOrbToolbar) = APITRUE Then
        MsgBox "Remove the 'Start' Toolbar from your taskbar first." & vbCrLf & "Right click taskbar, Toolbars > 'Start'", vbCritical

        Exit Sub

    End If

    frmStartButton.ActivateMoveMode
End Sub

Sub ResetPosition()

    If IsWindow(TaskbarHelper.g_StartButtonHwnd) = APIFALSE Then
        If IsWindow(TaskbarHelper.g_viOrbToolbar) = APIFALSE Then
            frmInstall.Show vbModal
        End If

        If IsWindow(TaskbarHelper.g_viOrbToolbar) = APIFALSE Then

            Exit Sub

        End If
        
        frmZOrderKeeper.HijackZOrder
    End If

    frmStartButton.ResetPosition
End Sub

Private Sub Form_Initialize()
    Set m_PopupSystemMenu = New clsMenu

    m_PopupSystemMenu.AddItem 1, "&Exit"
    m_PopupSystemMenu.AddSeperater
    m_PopupSystemMenu.AddItem 4, "&Move"
    m_PopupSystemMenu.AddItem 5, "&Reset"
    m_PopupSystemMenu.AddSeperater
    m_PopupSystemMenu.AddItem 6, "&Pick New Orb Image"
    m_PopupSystemMenu.AddItem 3, "&Settings"
    m_PopupSystemMenu.AddItem 2, "&About"

End Sub

Private Sub Form_Load()
    
    Dim theTip()    As Byte: theTip = App.Title & vbNullChar

    Dim theTipIndex As Long
    
    With nid
        .cbSize = Len(nid)
        .hWnd = Me.hWnd
        .uID = App.hInstance
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
        .HICON = Me.Icon
     
        For theTipIndex = 0 To UBound(theTip)
            .szTip(theTipIndex) = theTip(theTipIndex)
        Next

    End With
    
    ShowWindow Me.hWnd, SW_HIDE
    Shell_NotifyIcon NIM_ADD, nid

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    'this procedure receives the callbacks from the System Tray icon.
    Dim msg As Long
 
    'the value of X will vary depending upon the scalemode setting
    If Me.ScaleMode = vbPixels Then
        msg = x
    Else
        msg = x / Screen.TwipsPerPixelX
    End If
    
    Select Case msg
      
        Case WM_RBUTTONUP        '517 display popup menu
            SetForegroundWindow Me.hWnd
            PopupSystemMenu

    End Select

End Sub

Private Sub Form_Resize()

    On Error Resume Next
    
    'this is necessary to assure that the minimized window is hidden
    If Me.WindowState = vbMinimized Then Me.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'this removes the icon from the system tray
    Shell_NotifyIcon NIM_DELETE, nid
End Sub

Private Sub timExit_Timer()
    On Error Resume Next

    timExit.Enabled = False
    
    Unload frmClock
    Unload frmStartButton
    UnloadApplication
End Sub
