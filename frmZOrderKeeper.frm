VERSION 5.00
Begin VB.Form frmZOrderKeeper 
   Caption         =   "Container"
   ClientHeight    =   3315
   ClientLeft      =   -76680
   ClientTop       =   17415
   ClientWidth     =   3345
   LinkTopic       =   "Form1"
   ScaleHeight     =   3315
   ScaleWidth      =   3345
   ShowInTaskbar   =   0   'False
   Begin VB.Timer timZOrderChecker 
      Interval        =   500
      Left            =   480
      Top             =   2400
   End
End
Attribute VB_Name = "frmZOrderKeeper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'    Component  : frmZOrderKeeper
'    Project    : ViOrb5
'
'    Description: Keeps our windows above taskbar (but not over full screen apps)
'    Tricks: When this form is owned by the windows taskbar, this form
'    gets activated on top the taskbar whenever the taskbar is activated
'    (SetOwner)
'
'
'    Modified   :
'--------------------------------------------------------------------------------

Option Explicit

Private m_NotTopMost As Boolean


'--------------------------------------------------------------------------------
' Procedure  :       HijackZOrder
' Description:       Pretends to be a child window of the official taskbar or
'                    start button (where available) to keep same zorder
' Parameters :
'--------------------------------------------------------------------------------
Sub HijackZOrder()
    '    Exit Sub
    
    TaskbarHelper.UpdatehWnds

    frmFader.Show
    frmStartButton.Show

    'makesure our start button is owned by this window (for keeping zOrder)
    SetOwner frmFader.hWnd, Me.hWnd
    SetOwner frmStartButton.hWnd, Me.hWnd

    'fake start button placeholder is not available
    If IsWindow(TaskbarHelper.g_viOrbToolbar) = APIFALSE Then
        
        If g_Windows81 And Not frmStartButton.ManualPosition Then

            frmStartButton.Mode = NORMAL_MODE
            
            'make the real windows taskbar our parent, thus forcing
            'us above all other desktop windows naturally (like the taskbar)
            SetParent frmStartButton.hWnd, TaskbarHelper.g_TaskBarHwnd
                        
        'running Windows 8.1/or an OS with a REAL START button
        Else
            frmStartButton.Mode = LAYERED_MODE 'float as usual (Vista Orb style)
            
            'make the real start button our parent, keeping same zorder with it
            SetParent Me.hWnd, TaskbarHelper.g_StartButtonHwnd
        End If
        
        ShowWindow g_StartButtonHwnd, SW_HIDE 'hide any valid start button
                                              'handle we found
    
    'Our fake start button placeholder is available
    Else

        If IsWindow(TaskbarHelper.g_viOrbToolbar) = APITRUE Then '<< not nessarcy
            ViOrbToolbar = True
            
            frmStartButton.Mode = NORMAL_MODE
            SetParent frmStartButton.hWnd, TaskbarHelper.g_ReBarWindow32Hwnd
        Else
        
            'this code should never get reached
            'TODO: Remove unnessarcy unreachable code
            SetParent frmStartButton.hWnd, TaskbarHelper.g_TaskBarHwnd
        End If
    End If
    
    frmStartButton.Show
    frmStartButton.Move 0, 0
    
     'keep fader animation window above other windows globally, regardless
     '(it should be fine for an invisible window to always be on top of
     'everything else)
    StayOnTop frmFader, True
    
    If frmStartButton.Mode = LAYERED_MODE Then
        StayOnTop frmStartButton, True
        frmFader.Show vbModeless, frmStartButton
    End If
    
    'frmFader.Show vbModeless, frmStartButton
    frmFader.timFadeIn.Enabled = True

    Me.Visible = False
End Sub

Private Sub Form_Load()
    HijackZOrder
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ShowWindow g_StartButtonHwnd, SW_SHOW
End Sub


'--------------------------------------------------------------------------------
' Procedure  :       timZOrderChecker_Timer
' Description:       Attempts to restore zOrder to choas (assumes choas)
' Parameters :
'--------------------------------------------------------------------------------
Private Sub timZOrderChecker_Timer()

    'Enforces Z-Order's
    Dim hWndForeGroundWindow As Long

    Dim zOrderTaskBar        As Long
    
    hWndForeGroundWindow = GetForegroundWindow
    zOrderTaskBar = GetZOrder(frmStartButton.hWnd)

    If (GetZOrder(g_TaskBarHwnd) < zOrderTaskBar) And m_NotTopMost = False Then

        SetWindowPos frmStartButton.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE
    Else

        'if the foreground window isn't one of our own windows and it isn't the windows taskbar or the windows start menu -then
        If Not hWndBelongToUs(hWndForeGroundWindow) And _
        hWndForeGroundWindow <> TaskbarHelper.g_TaskBarHwnd And _
        hWndForeGroundWindow <> TaskbarHelper.g_StartMenuHwnd Then
        
           'if the foreground window is behind the taskbar -then
            If IsTaskBarBehindWindow(hWndForeGroundWindow) Then
                'if the foreground window is not a topmost window
                If IsWindowTopMost(hWndForeGroundWindow) = False Then
                
                    'theoritically ViOrb might be appearing in front on top of
                    'the foreground window (that's behind the taskbar) so it's
                    'a full screen window ontop the taskbar
                    Debug.Print "Hiding ViOrb"

                    m_NotTopMost = True

                    frmStartButton.Hide
                    frmFader.Hide
                Else

                    If zOrderTaskBar < GetZOrder(hWndForeGroundWindow) Then

                        m_NotTopMost = True 'viglance/vistart stealing focus?
                        SetWindowPos hWndForeGroundWindow, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
                    End If
                End If
                
                'Me.Hide
            Else

                If m_NotTopMost = True Then
                    m_NotTopMost = False
                    
                    frmStartButton.Show
                    frmFader.Show
                End If
            End If
        End If
    End If

    Exit Sub

Handler:
    LogError Err.Number, "zOrderCheck(" & Err.Description & ")", "winZOrder"
End Sub

