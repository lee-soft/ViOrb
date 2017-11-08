VERSION 5.00
Begin VB.Form frmStartButton 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "#Start~ViOrb#"
   ClientHeight    =   1305
   ClientLeft      =   3315
   ClientTop       =   3750
   ClientWidth     =   1485
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   87
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   99
   ShowInTaskbar   =   0   'False
   Tag             =   "Start"
   Begin VB.Timer timKeepOnStartButton 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer timFollowCursor 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer timStartMenuCheck 
      Interval        =   100
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "frmStartButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'    Component  : frmStartButton
'    Project    : ViOrb5
'
'    Description: This actual start button object supports two rendering modes
'                 layered window mode:
' https://msdn.microsoft.com/en-us/library/windows/desktop/ms632599(v=vs.85).aspx
'                 Or normal mode. Each mode supports different features.
'
'    Modified   :
'--------------------------------------------------------------------------------
Option Explicit

Private ORB_HEIGHT               As Long

Private ORB_WIDTH                As Long

Private m_Position               As POINTL

Private m_nudge                  As POINTL

Private m_manualPosition         As Boolean

Private m_theStartButton         As GDIPImage

Private WithEvents m_FaderWindow As frmFader
Attribute m_FaderWindow.VB_VarHelpID = -1

' Create a Graphics object:
Private m_gfx                    As GDIPGraphics

Private m_Bitmap                 As GDIPBitmap

Private m_BitmapGraphics         As GDIPGraphics

Private m_SourcePositionY        As Long

Private m_reBar32_Rect           As RECT

Private m_Rect                   As RECT

Private m_startbuttonFileName    As String

Private m_mode                   As PNG_DRAWMODE

Private m_win78                  As GDIPImage

Private m_layeredWindow          As LayerdWindowHandles


'--------------------------------------------------------------------------------
' Procedure  :       ActivateMoveMode
' Description:       makes the start button follow the user's
'                    mouse cursor.
' Parameters :
'--------------------------------------------------------------------------------
Sub ActivateMoveMode()
    m_manualPosition = True

    If m_mode = NORMAL_MODE Then
        frmZOrderKeeper.HijackZOrder
        frmOptions.ForceRefresh
        frmOptions.timDelayMoveMode.Enabled = True

        Exit Sub

    End If

    timFollowCursor.Enabled = True
End Sub


'--------------------------------------------------------------------------------
' Procedure  :       MoveOrbIfNotOverStartButton
' Description:       Repositions 'this' form over the "start button"
'                    area for any supported windows OS
' Parameters :
'--------------------------------------------------------------------------------
Public Function MoveOrbIfNotOverStartButton()

    If g_Windows81 Then
        MoveOrbOverWindows8StartButton

        Exit Function

    End If
    
    If ViOrbToolbar Then
        MoveOrbViOrbToolbar TaskbarHelper.g_viOrbToolbar

        Exit Function

    End If

    Dim recStartButton As RECT

    Dim lngTop         As Long

    Dim finalX         As Long

    Dim finalY         As Long

    GetWindowRect g_ReBarWindow32Hwnd, m_reBar32_Rect
    GetWindowRect g_StartButtonHwnd, recStartButton
    GetWindowRect Me.hWnd, m_Rect


    If GetTaskBarEdge = ABE_LEFT Or GetTaskBarEdge = ABE_RIGHT Then

        lngTop = -1
    Else

        If m_reBar32_Rect.bottom - m_reBar32_Rect.Top < 40 Then
            
            lngTop = 12
        Else
            lngTop = (ORB_HEIGHT / 2) - (m_reBar32_Rect.bottom - m_reBar32_Rect.Top) / 2
        End If
    End If

    If lngTop <> -1 Then
        finalX = (recStartButton.Left) + m_nudge.x
        finalY = (m_reBar32_Rect.Top - lngTop) + m_nudge.y
    
        If finalX <> (m_Rect.Left) Or (finalY <> m_Rect.Top) Then
    
            MoveWindow Me.hWnd, finalX, finalY, Me.ScaleWidth, Me.ScaleHeight, False
            SnapFaderOverMe
        End If

    Else
        finalX = (recStartButton.Left) + m_nudge.x
        finalY = (recStartButton.Top) + m_nudge.y
        
        If finalX <> (m_Rect.Left) Or (finalY <> m_Rect.Top) Then
            
            MoveWindow Me.hWnd, finalX, finalY, Me.ScaleWidth, Me.ScaleHeight, False
            SnapFaderOverMe
        End If
    End If

End Function


'--------------------------------------------------------------------------------
' Procedure  :       MoveOrbOverWindows8StartButton
' Description:       Try to aproximate the start button position
'                    TODO: find a more precise way to estimate the position
' Parameters :
'--------------------------------------------------------------------------------
Public Function MoveOrbOverWindows8StartButton()

    Dim taskbarEdge     As AbeBarEnum

    Dim taskBarHeight   As Long

    Dim taskbarWidth    As Long

    Dim finalX          As Long

    Dim finalY          As Long

    taskbarEdge = GetTaskBarEdge()
    finalX = m_nudge.x
    finalY = m_nudge.y
        
    GetWindowRect g_ReBarWindow32Hwnd, m_reBar32_Rect

    taskBarHeight = (m_reBar32_Rect.bottom - m_reBar32_Rect.Top)
    taskbarWidth = (m_reBar32_Rect.Right - m_reBar32_Rect.Left)
    
    If taskbarEdge = ABE_RIGHT Or taskbarEdge = ABE_LEFT Then
 
        finalX = finalX + (taskbarWidth / 2) - (Me.ScaleHeight / 2)

    ElseIf taskbarEdge = abe_bottom Or taskbarEdge = ABE_TOP Then
    
        finalY = finalY + ((((ORB_HEIGHT) / 2) - (taskBarHeight) / 2) * -1) + 2
    
    End If
    
    MoveWindow Me.hWnd, finalX, finalY, Me.ScaleWidth, Me.ScaleHeight, 1
    SnapFaderOverMe
End Function


'--------------------------------------------------------------------------------
' Procedure  :       MoveOrbViOrbToolbar
' Description:       Repositions the start button over the start button
'                    placeholder
' Parameters :
'--------------------------------------------------------------------------------
Public Function MoveOrbViOrbToolbar()

    Dim taskbarEdge     As AbeBarEnum

    Dim taskBarHeight   As Long

    Dim recViOrbToolbar As RECT
     
    Dim finalX          As Long

    Dim finalY          As Long

    taskbarEdge = GetTaskBarEdge()
    finalX = m_nudge.x
    finalY = m_nudge.y
        
    GetWindowRect g_ReBarWindow32Hwnd, m_reBar32_Rect
    GetClientRect TaskbarHelper.g_viOrbToolbar, recViOrbToolbar
        
    taskBarHeight = (m_reBar32_Rect.bottom - m_reBar32_Rect.Top)

    If taskbarEdge = ABE_RIGHT Or taskbarEdge = ABE_LEFT Then
 
        finalX = finalX + (taskBarHeight / 2) - (Me.ScaleHeight / 2)

    ElseIf taskbarEdge = abe_bottom Or taskbarEdge = ABE_TOP Then
    
        finalY = finalY + ((((ORB_HEIGHT) / 2) - (taskBarHeight) / 2) * -1) + 2
    
    End If
    
    MoveWindow Me.hWnd, finalX, finalY, Me.ScaleWidth, Me.ScaleHeight, 1
    SnapFaderOverMe
End Function


'--------------------------------------------------------------------------------
' Procedure  :       ResetPosition
' Description:       Reset's start button physical position back to defaults
'                    removing manually assigned user position
' Parameters :
'--------------------------------------------------------------------------------
Sub ResetPosition()
    m_manualPosition = False
    
    If g_Windows81 Then
        frmOptions.timDelayRefresh = True

        Exit Sub

    End If

    If IsWindow(TaskbarHelper.g_viOrbToolbar) = APITRUE Then
        ViOrbToolbar = True
    Else
        ViOrbToolbar = False
    End If
    
    RegistryHelper.DeleteKey AppSettingsRegistryPath & "lastX"
    
    If frmOptions.Visible Then frmOptions.UpdateCaptions
End Sub


'--------------------------------------------------------------------------------
' Procedure  :       UpdateAndReDraw
' Description:       Clears the surface and redraws the button
' Parameters :       UpdateHDC (Boolean = True)
'--------------------------------------------------------------------------------
Sub UpdateAndReDraw(Optional ByVal UpdateHDC As Boolean = True)

    m_BitmapGraphics.Clear
    m_BitmapGraphics.DrawImageRect m_theStartButton, 0, 0, ORB_WIDTH, ORB_HEIGHT, 0, m_SourcePositionY

    m_gfx.Clear

    If m_mode = NORMAL_MODE Then
        'this makes the background appear transparent on the windows 7 taskbar,
        'TODO: check it works when the taskbar is a different tint colour
        m_gfx.DrawImageRectFv m_win78, 0, 0, Me.ScaleWidth * 2, Me.ScaleHeight * 2
    End If

    'regardless of the state we're in, we still need to render the start button
    m_gfx.DrawImageRectFv m_Bitmap.Image, 0, 0, Me.ScaleWidth, Me.ScaleHeight

    If m_mode = LAYERED_MODE Then
        If UpdateHDC Then m_layeredWindow.Update Me.hWnd, Me.hdc, 255
    Else
        Me.Refresh
    End If

End Sub

Private Function GetMyRect() As win.RECT
    GetWindowRect Me.hWnd, GetMyRect
End Function

Private Sub RetrieveStoredPosition()

    On Error GoTo Handler

    m_Position.x = -1
    m_Position.x = ReadKeyInteger(AppSettingsRegistryPath & "lastX", -1)
    m_Position.y = ReadKeyInteger(AppSettingsRegistryPath & "lastY", -1)

    If IsWindow(TaskbarHelper.g_viOrbToolbar) = APIFALSE And ReadKeyInteger(AppSettingsRegistryPath & "manual_position", 0) = 1 Then
        
        m_manualPosition = True
    End If

    If ReadKeyInteger(AppSettingsRegistryPath & "spoof_viglance", 1) = 1 Then
        Me.Caption = MainHelper.ViGlance_Identifier
    End If
    
    m_nudge.x = ReadKeyInteger(AppSettingsRegistryPath & GetFilenameFromPath(m_startbuttonFileName) & "\nudgeX", 0)
    m_nudge.y = ReadKeyInteger(AppSettingsRegistryPath & GetFilenameFromPath(m_startbuttonFileName) & "\nudgeY", 0)
    
    If m_Position.x > Screen.Width Then
        m_Position.x = -1
    End If
    
    If m_Position.y > Screen.Height Then
        m_Position.y = -1
    End If
    
    timKeepOnStartButton_Timer
Handler:
End Sub


'--------------------------------------------------------------------------------
' Procedure  :       SetFaderWindow
' Description:       The routine that intializes the fader window
' Parameters :
'--------------------------------------------------------------------------------
Private Function SetFaderWindow()
    'On Error Resume Next
    
    Set m_FaderWindow = frmFader
    Load m_FaderWindow
    
    m_FaderWindow.InitializeOrb m_theStartButton
    m_FaderWindow.FrameIndex = 1
    m_FaderWindow.Show
    
    SnapFaderOverMe
End Function


'--------------------------------------------------------------------------------
' Procedure  :       SnapFaderOverMe
' Description:       Reposition fading window over our start button
' Parameters :
'--------------------------------------------------------------------------------
Private Function SnapFaderOverMe()
    'm_FaderWindow.Move Me.Left, Me.Top

    GetWindowRect Me.hWnd, m_Rect
    SetWindowPos m_FaderWindow.hWnd, 0, m_Rect.Left, m_Rect.Top, 0, 0, SWP_NOSIZE Or SWP_NOREPOSITION Or SWP_NOACTIVATE
End Function

Private Sub Form_Initialize()

    m_startbuttonFileName = App.Path & "\Resources\" & ReadKeyString(AppSettingsRegistryPath & "filename")

    Set m_gfx = New GDIPGraphics
    Set m_BitmapGraphics = New GDIPGraphics
    Set m_Bitmap = New GDIPBitmap
    Set m_theStartButton = New GDIPImage
    Set m_win78 = New GDIPImage
    
    If FileExists(App.Path & "\orb_background.png") Then
        m_win78.FromFile App.Path & "\orb_background.png"
    Else
        'FromResource
        m_win78.FromStream LoadResData("WIN7", "PNG")
    End If
    
    If Not FileExists(m_startbuttonFileName) Then
        m_startbuttonFileName = App.Path & "\Resources\Windows 7 Orb.png"
    End If
    
    If Not FileExists(m_startbuttonFileName) Then UnloadApplication
    m_theStartButton.FromFile m_startbuttonFileName
    
    'since there's 3 start button states occupying the same space in our graphic
    'we can assume that the start button height is therefore 1 third our
    'graphic's height
    ORB_HEIGHT = m_theStartButton.Height / 3
    ORB_WIDTH = m_theStartButton.Width
    
    Me.Width = ORB_WIDTH * Screen.TwipsPerPixelX
    Me.Height = ORB_HEIGHT * Screen.TwipsPerPixelY
    
    RetrieveStoredPosition

    m_Bitmap.CreateFromSizeFormat ORB_WIDTH, ORB_HEIGHT, PixelFormat32bppARGB
    m_BitmapGraphics.FromImage m_Bitmap.Image

    SetFaderWindow
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' in a perfect world this window shouldn't recieve a mouse down event as
    ' the fader is always on top of it in case it does though, we just react
    ' the same way
    m_FaderWindow_onClicked
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'force the fader window above this window (so it does the fade effect)
    StayOnTop m_FaderWindow, True
End Sub

Private Sub Form_Resize()

    On Error Resume Next

    
    If m_mode = LAYERED_MODE Then

        'when a layered window is resized you need to recreate it
        'though really this should only happen once per skin change (new graphic)
        Set m_layeredWindow = Nothing
        Set m_layeredWindow = MakeLayerdWindow(Me)
        
        m_gfx.FromHDC m_layeredWindow.theDC
        
    ElseIf m_mode = NORMAL_MODE Then

        'the form's hDC will be invalid, recreate
        m_gfx.FromHDC Me.hdc
    End If
    
    UpdateAndReDraw True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'dump all the settings to registry

    If m_manualPosition Then
        WriteRegistryInteger AppSettingsRegistryPath & "lastX", m_Position.x
        WriteRegistryInteger AppSettingsRegistryPath & "lastY", m_Position.y
        WriteRegistryInteger AppSettingsRegistryPath & "manual_position", 1
    Else
        WriteRegistryInteger AppSettingsRegistryPath & "manual_position", 0
    End If
    
    If Me.Caption = MainHelper.ViGlance_Identifier Then
        WriteRegistryInteger AppSettingsRegistryPath & "spoof_viglance", 1
    Else
        WriteRegistryInteger AppSettingsRegistryPath & "spoof_viglance", 0
    End If
    
    WriteRegistryInteger AppSettingsRegistryPath & GetFilenameFromPath(m_startbuttonFileName) & "\nudgeY", m_nudge.y
    WriteRegistryInteger AppSettingsRegistryPath & GetFilenameFromPath(m_startbuttonFileName) & "\nudgeX", m_nudge.x
                
    Unload m_FaderWindow
    Set m_FaderWindow = Nothing
    
    m_gfx.Dispose
    m_Bitmap.Dispose
    m_theStartButton.Dispose
    m_BitmapGraphics.Dispose
    
    DisposeGDIIfLast
End Sub

'--------------------------------------------------------------------------------
' Procedure  :       m_FaderWindow_onClicked
' Description:       [type_description_here]
' Parameters :
'--------------------------------------------------------------------------------
Private Sub m_FaderWindow_onClicked()

    Const WINKEY = 91

    If timFollowCursor.Enabled Then
    
        m_manualPosition = True
        m_Position.x = Me.Left / Screen.TwipsPerPixelX
        m_Position.y = Me.Top / Screen.TwipsPerPixelY
    
        timFollowCursor.Enabled = False

        If frmOptions.Visible Then frmOptions.UpdateCaptions
        
        Exit Sub

    End If
    
    If g_StartMenuOpen = False Then
        If g_viStartRunning = False Then
            
            TaskbarHelper.ShowStartMenu
        Else
            
            'since vistart has hooked the windows key this should summon it
            SetKeyDown WINKEY
            SetKeyUp WINKEY
            
        End If
    End If

End Sub

Private Sub m_FaderWindow_onMouseUp(Button As Integer)

    If Button = vbRightButton Then
        frmClock.PopupSystemMenu
    End If

End Sub

Private Sub m_FaderWindow_onRolledOut()

    If Not IsStartMenuOpen Then
        FrameIndex = 0
    
        UpdateAndReDraw False
        m_FaderWindow.FadeOut
    End If

End Sub

Private Sub m_FaderWindow_onRolledOver()

    If Not IsStartMenuOpen Then
        
        'frameIndex = 1
        
        m_FaderWindow.Alpha = 1
        m_FaderWindow.FrameIndex = 1
        m_FaderWindow.UpdateAndReDraw True
        
        UpdateAndReDraw False
        
        m_FaderWindow.FadeIn
        m_FaderWindow.ZOrder 0
    End If

End Sub


'--------------------------------------------------------------------------------
' Procedure  :       timFollowCursor_Timer
' Description:       Repositons our start button so it appears where the mouse
'                    cursor is
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub timFollowCursor_Timer()

    Dim cPos As win.POINTL

    GetCursorPos cPos
    
    MoveWindow Me.hWnd, cPos.x - (Me.ScaleWidth / 2), cPos.y - (Me.ScaleHeight / 2), Me.ScaleWidth, Me.ScaleHeight, 0
                                                 
    SnapFaderOverMe
End Sub


'--------------------------------------------------------------------------------
' Procedure  :       timKeepOnStartButton_Timer
' Description:       Keeps the fading overlay aligned on top of the start button
' Parameters :
'--------------------------------------------------------------------------------
Private Sub timKeepOnStartButton_Timer()

    If timFollowCursor.Enabled Then Exit Sub

    GetWindowRect Me.hWnd, m_Rect
    
    If m_manualPosition Then
        If m_Rect.Left <> m_Position.x Or m_Rect.Top <> m_Position.y Then
            MoveWindow Me.hWnd, m_Position.x, m_Position.y, ORB_WIDTH, ORB_HEIGHT, 0
            SnapFaderOverMe
        End If

    Else
        MoveOrbIfNotOverStartButton
    End If

End Sub


'--------------------------------------------------------------------------------
' Procedure  :       timStartMenuCheck_Timer
' Description:       periodically polls the start menu/vistart menu state
'                    to check if it's open or closed then sets the state of
'                    the start button accordingly
' Parameters :
'--------------------------------------------------------------------------------
Private Sub timStartMenuCheck_Timer()
    g_StartMenuOpen = IsStartMenuOpen

    If IsStartMenuOpen Then
        If m_FaderWindow.FrameIndex <> 2 Then

            m_FaderWindow.FrameIndex = 2
            m_FaderWindow.UpdateAndReDraw True
        End If
        
    ElseIf Not IsStartMenuOpen Then

        If m_FaderWindow.FrameIndex = 2 Then
            m_FaderWindow.FrameIndex = 1
            m_FaderWindow.UpdateAndReDraw True
            
            If Not PointInsideOfRect(GetCursorPoint(), GetMyRect()) Then
                Debug.Print "INSIDE!"
                m_FaderWindow.FadeOut
            End If

        Else

            If m_FaderWindow.FrameIndex <> 1 Then
                m_FaderWindow.FrameIndex = 1
                m_FaderWindow.UpdateAndReDraw True
            End If
        End If
    End If

End Sub

Public Property Get ManualPosition() As Boolean
    ManualPosition = m_manualPosition
End Property

Public Property Get Mode() As PNG_DRAWMODE
    Mode = m_mode
End Property

'--------------------------------------------------------------------------------
' Procedure  :       mode
' Description:       Attempts to set a new rendering mode on the fly without the
'                    need to reload the entire window or application
' Parameters :       newMode (PNG_DRAWMODE)
'--------------------------------------------------------------------------------
Public Property Let Mode(newMode As PNG_DRAWMODE)
    m_mode = CLng(newMode)
    
    If m_mode = NORMAL_MODE Then
        Me.AutoRedraw = True
    Else
        Me.AutoRedraw = False
        
        If Not m_layeredWindow Is Nothing Then
            m_layeredWindow.Release
            Set m_layeredWindow = Nothing
        End If
    End If
    
    Form_Resize
End Property

'--------------------------------------------------------------------------------
' Procedure  :       NudgeX
' Description:       [type_description_here]
' Parameters :
'--------------------------------------------------------------------------------
Public Property Get NudgeX() As Long
    NudgeX = m_nudge.x
End Property


'--------------------------------------------------------------------------------
' Procedure  :       NudgeX
' Description:       Nudges start button position to new X cordinate
' Parameters :       newX (Long)
'--------------------------------------------------------------------------------
Public Property Let NudgeX(newX As Long)
    m_nudge.x = newX
    timKeepOnStartButton_Timer 'to see changes immediately
End Property

'--------------------------------------------------------------------------------
' Procedure  :       NudgeY
' Description:       [type_description_here]
' Parameters :
'--------------------------------------------------------------------------------
Public Property Get NudgeY() As Long
    NudgeY = m_nudge.y
End Property


'--------------------------------------------------------------------------------
' Procedure  :       NudgeY
' Description:       Nudges start button to new Y cordinate
' Parameters :       newY (Long)
'--------------------------------------------------------------------------------
Public Property Let NudgeY(newY As Long)

    m_nudge.y = newY
    timKeepOnStartButton_Timer 'to see changes immediately
    
End Property


'--------------------------------------------------------------------------------
' Procedure  :       FrameIndex
' Description:       Each start button state is represented by an integer(1-3)
'                    this lets you easily switch buttons
' Parameters :       newIndex (Long)
'--------------------------------------------------------------------------------
Private Property Let FrameIndex(ByVal newIndex As Long)
    m_SourcePositionY = newIndex * ORB_HEIGHT
End Property
