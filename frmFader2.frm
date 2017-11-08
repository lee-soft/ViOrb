VERSION 5.00
Begin VB.Form frmFader 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   0  'None
   Caption         =   "Start_Fader"
   ClientHeight    =   4185
   ClientLeft      =   10170
   ClientTop       =   5235
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   ScaleHeight     =   279
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   360
   ShowInTaskbar   =   0   'False
   Tag             =   "Start"
   Begin VB.Timer timFadeIn 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer timFadeOut 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   480
      Top             =   0
   End
End
Attribute VB_Name = "frmFader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'    Component  : frmFader
'    Project    : ViOrb5
'
'    Description: Creates the fading effect when you mouse over the start button
'                 by always sitting on top of the start button. Unlike the
'                 frmStartButton window, this Window is ALWAYS a layered window
'    Modified   :
'--------------------------------------------------------------------------------
Option Explicit

Public Event onRolledOver() 'Fires when the cursor enters this window

Public Event onRolledOut()  'Fires when the cursor leaves this window

Public Event onClicked()    'Fires when the mouse is clicked on this window

Public Event onMouseUp(Button As Integer) 'Fires when the mouse button has been
                                          'released

Private ORB_HEIGHT          As Long

Private ORB_WIDTH           As Long

Private m_theStartButton    As GDIPImage 'contains the start button graphic

Private m_MouseInClientArea As Boolean

Private m_MouseEvents       As TrackMouseEvent

Private m_Alpha             As Byte 'the current alpha value for fading

Private lastAlpha           As Long

Private m_gfx               As GDIPGraphics

Private m_Bitmap            As GDIPBitmap

Private m_BitmapGraphics    As GDIPGraphics

Private m_SourcePositionY   As Long

Private m_introDone         As Boolean

Private m_layeredWindow     As LayerdWindowHandles

Implements IHookSink

Sub FadeIn()
    timFadeOut.Enabled = False
    timFadeIn.Enabled = True

End Sub

Sub FadeOut()
    timFadeOut.Enabled = True
    timFadeIn.Enabled = False
End Sub

Public Sub InitializeOrb(ByRef srcStartButton As GDIPImage)

    If InitializeGDIIfNotInitialized = False Then

        Exit Sub

    End If
    
    Set m_gfx = New GDIPGraphics
    Set m_BitmapGraphics = New GDIPGraphics
    Set m_Bitmap = New GDIPBitmap
    
    Set m_theStartButton = srcStartButton
    
    ORB_HEIGHT = m_theStartButton.Height / 3
    ORB_WIDTH = m_theStartButton.Width
    
    MoveWindow Me.hWnd, 0, 0, ORB_WIDTH, ORB_HEIGHT, 0
    
    m_Bitmap.CreateFromSizeFormat ORB_WIDTH, ORB_HEIGHT, PixelFormat32bppARGB
    m_BitmapGraphics.FromImage m_Bitmap.Image

    Set m_layeredWindow = Nothing
    Set m_layeredWindow = MakeLayerdWindow(Me)
    m_gfx.FromHDC m_layeredWindow.theDC
    UpdateAndReDraw

End Sub


'--------------------------------------------------------------------------------
' Procedure  :       UpdateAndReDraw
' Description:       Updates the bitmap and then re-draws it to the DC
' Parameters :       UpdateHDC (Boolean = True)
'--------------------------------------------------------------------------------
Sub UpdateAndReDraw(Optional ByVal UpdateHDC As Boolean = True)

    m_BitmapGraphics.Clear
    m_BitmapGraphics.DrawImageRect m_theStartButton, 0, 0, ORB_WIDTH, _
                                               ORB_HEIGHT, 0, m_SourcePositionY

    m_gfx.Clear
    m_gfx.DrawImageRectFv m_Bitmap.Image, 0, 0, Me.ScaleWidth, Me.ScaleHeight
    
    If UpdateHDC Then m_layeredWindow.Update Me.hWnd, Me.hdc, m_Alpha
End Sub

Private Function IHookSink_WindowProc(hWnd As Long, _
                                      iMsg As Long, _
                                      wParam As Long, _
                                      lParam As Long) As Long

    On Error GoTo Handler

    If Not m_introDone Then GoTo Handler
    
    If iMsg = WM_MOUSEMOVE Then
        If Not m_MouseInClientArea Then
            m_MouseInClientArea = True

            m_MouseEvents.cbSize = Len(m_MouseEvents)
            m_MouseEvents.dwFlags = TME_LEAVE
            m_MouseEvents.hwndTrack = Me.hWnd
            
            TrackMouseEvent m_MouseEvents
            RaiseEvent onRolledOver
            
            'GoTo Handler
        End If
        
    ElseIf iMsg = WM_MOUSELEAVE Then
        m_MouseInClientArea = False
        
        RaiseEvent onRolledOut
    ElseIf iMsg = WM_LBUTTONDOWN Then
        RaiseEvent onClicked
    Else
        ' Just allow default processing for everything else.
        IHookSink_WindowProc = InvokeWindowProc(hWnd, iMsg, wParam, lParam)
    End If
    
    Exit Function

Handler:
    ' Just allow default processing for everything else.
    IHookSink_WindowProc = InvokeWindowProc(hWnd, iMsg, wParam, lParam)
End Function

Private Sub Form_Initialize()
    Call HookWindow(Me.hWnd, Me)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbLeftButton Then
        ReleaseCapture
        Call SendMessage(ByVal Me.hWnd, ByVal WM_NCLBUTTONDOWN, ByVal HTCAPTION, 0&)
    End If

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent onMouseUp(Button)
End Sub

Private Sub Form_Terminate()
    UnhookWindow Me.hWnd
End Sub

Private Sub Form_Unload(Cancel As Integer)
    m_gfx.Dispose
    m_Bitmap.Dispose
    m_BitmapGraphics.Dispose
    
    DisposeGDIIfLast
    
    UnhookWindow Me.hWnd
End Sub


'--------------------------------------------------------------------------------
' Procedure  :       timFadeIn_Timer
' Description:       Fades the start button in
' Parameters :
'--------------------------------------------------------------------------------
Private Sub timFadeIn_Timer()

    On Error Resume Next

    If lastAlpha < 255 Then
        lastAlpha = lastAlpha + 15
        Alpha = CByte(lastAlpha)
    Else
        timFadeIn.Enabled = False
        
        If Not m_introDone Then
            m_introDone = True
            timFadeOut.Enabled = True
        End If
    End If

End Sub


'--------------------------------------------------------------------------------
' Procedure  :       timFadeOut_Timer
' Description:       Fades the start button out
' Parameters :
'--------------------------------------------------------------------------------
Private Sub timFadeOut_Timer()

    On Error Resume Next

    If lastAlpha > 1 Then
        lastAlpha = lastAlpha - 15
        Alpha = CByte(lastAlpha)
    Else
        timFadeOut.Enabled = False
        Alpha = CByte(1)
    End If

End Sub

Public Property Let Alpha(newAlpha As Byte)
    m_Alpha = newAlpha
    UpdateAndReDraw True
End Property

Property Get FrameIndex() As Long
    FrameIndex = m_SourcePositionY / ORB_HEIGHT
End Property

Property Let FrameIndex(ByVal newIndex As Long)
    m_SourcePositionY = newIndex * ORB_HEIGHT
End Property


