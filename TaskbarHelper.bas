Attribute VB_Name = "TaskbarHelper"
'--------------------------------------------------------------------------------
'    Component  : TaskbarHelper
'    Project    : ViOrb5
'
'    Description: This module contains all logic involving the Windows
'                 taskbar.
'
'    Modified   :
'--------------------------------------------------------------------------------

Option Explicit

Public g_ReBarWindow32Hwnd As Long

Public g_StartButtonHwnd   As Long

Public g_TaskBarHwnd       As Long

Public g_StartMenuHwnd     As Long

Public g_StartMenuOpen     As Boolean

Public g_viStartRunning    As Boolean

'Public g_viStartOrbHwnd    As Long

Public g_viOrbToolbar      As Long


'--------------------------------------------------------------------------------
' Procedure  :       GetTaskBarEdge
' Description:       Determines the orientation of the taskbar, LEFT/RIGHT etc
' Parameters :
'--------------------------------------------------------------------------------
Function GetTaskBarEdge() As AbeBarEnum
        
    Dim abd As APPBARDATA

    abd.cbSize = LenB(abd)
    abd.hWnd = g_TaskBarHwnd
    SHAppBarMessage ABM_GETTASKBARPOS, abd
    
    GetTaskBarEdge = GetEdge(abd.rc)

End Function

'--------------------------------------------------------------------------------
' Procedure  :       IsStartMenuOpen
' Description:       Determines if the windows start menu or vistart is open
'                    would be nice to include other 3rd party start menus here
' Parameters :
'--------------------------------------------------------------------------------
Public Function IsStartMenuOpen() As Boolean

    If IsViStartOpen Then
        IsStartMenuOpen = True

        Exit Function

    End If

    If IsWindow(g_StartMenuHwnd) = False Then
        g_StartMenuHwnd = FindWindow("DV2ControlHost", "Start Menu")

        If g_StartMenuHwnd = 0 Then
            g_StartMenuHwnd = FindWindow("DV2ControlHost", vbNullString)
        End If
    End If
    
    IsStartMenuOpen = IsWindowVisible(g_StartMenuHwnd)
End Function


'--------------------------------------------------------------------------------
' Procedure  :       IsTaskBarBehindWindow
' Description:       Attempts to determine if a given Window handle is behind
'                    the Windows taskbar
' Parameters :       hWnd (Long)
'--------------------------------------------------------------------------------
Function IsTaskBarBehindWindow(hWnd As Long)
    
    If GetZOrder(g_TaskBarHwnd) > GetZOrder(hWnd) Then
        IsTaskBarBehindWindow = True
    Else
        IsTaskBarBehindWindow = False
    End If
    
End Function


'--------------------------------------------------------------------------------
' Procedure  :       IsViStartOpen
' Description:       Determines if ViStart is open specifically
' Parameters :
'--------------------------------------------------------------------------------
Public Function IsViStartOpen()

    IsViStartOpen = False
    
    If FindWindow("ThunderRT6FormDC", "ViStart_PngNew") <> 0 Then
        IsViStartOpen = True
    End If
    
    'Debug.Print "VGMODE::" & IsViStartOpen
    
End Function


'--------------------------------------------------------------------------------
' Procedure  :       IsWindowTopMost
' Description:       Determines if the given hWnd has the TOPMOST style
' Parameters :       hWnd (Long)
'--------------------------------------------------------------------------------
Function IsWindowTopMost(hWnd As Long)

    Dim windowStyle As Long

    IsWindowTopMost = False
    windowStyle = GetWindowLong(hWnd, GWL_EXSTYLE)
    
    If IsStyle(windowStyle, WS_EX_TOPMOST) Then
        IsWindowTopMost = True
    End If

End Function


'--------------------------------------------------------------------------------
' Procedure  :       ShowStartMenu
' Description:       Triggers the window start menu (ViStart can react to this too)
' Parameters :
'--------------------------------------------------------------------------------
Public Function ShowStartMenu()
    SendMessage g_TaskBarHwnd, ByVal WM_SYSCOMMAND, ByVal SC_TASKLIST, ByVal 0
End Function

'--------------------------------------------------------------------------------
' Procedure  :       UpdatehWnds
' Description:       Gets all the Window handles associated with the TaskBar
' Parameters :
'--------------------------------------------------------------------------------
Public Function UpdatehWnds() As Boolean

    Dim newTaskBarHwnd As Long

    Dim updatedHwnd    As Boolean

    Dim lParamReturn   As Long

    updatedHwnd = False
    
    newTaskBarHwnd = FindWindow("Shell_TrayWnd", "")

    If newTaskBarHwnd = 0 Then

        Exit Function

    End If

    If newTaskBarHwnd <> g_TaskBarHwnd Then
        updatedHwnd = True
    
        g_TaskBarHwnd = newTaskBarHwnd
        g_ReBarWindow32Hwnd = FindWindowEx(ByVal g_TaskBarHwnd, ByVal 0&, "ReBarWindow32", vbNullString)
        g_viOrbToolbar = FindWindowEx(ByVal g_ReBarWindow32Hwnd, ByVal 0&, "ToolbarWindow32", "Start")
        
        If g_WindowsXP = True Then
            g_StartButtonHwnd = FindWindowEx(g_TaskBarHwnd, 0, "Button", vbNullString)

            If g_StartButtonHwnd = 0 Then
                'Reset update trigger (forcing routine to later update again)
                g_TaskBarHwnd = -1
            End If

        Else
            g_StartButtonHwnd = FindWindowEx(g_TaskBarHwnd, 0, "Button", vbNullString)
            
            If g_StartButtonHwnd = 0 Then
                g_StartButtonHwnd = FindWindow("Button", "Start")

                If g_StartButtonHwnd = 0 Then g_StartButtonHwnd = FindWindow("Button", vbNullString)
                
                If g_StartButtonHwnd = 0 Then
                    Call EnumChildWindows(g_TaskBarHwnd, AddressOf EnumTaskbarChildrenToFindStartButton, lParamReturn)
                End If
                
            End If
            
            If g_StartButtonHwnd = 0 Then

                'Reset update trigger (forcing routine to later update again)
                If g_TaskBarHwnd > 0 And Not g_Windows8 Then g_TaskBarHwnd = -1
            End If
        End If
        
    End If
    
    UpdatehWnds = updatedHwnd
End Function


'--------------------------------------------------------------------------------
' Procedure  :       WaitForTaskbar
' Description:       A crude way to suspend program thread until
'                    we have the windows explorer's taskbar
' Parameters :
'--------------------------------------------------------------------------------
Public Function WaitForTaskbar() As Boolean

    Dim findAttempts As Long

    UpdatehWnds

    While IsWindow(g_TaskBarHwnd) = APIFALSE And findAttempts < 10

        findAttempts = findAttempts + 1
        
        Sleep 1000
        UpdatehWnds
        DoEvents

    Wend

    WaitForTaskbar = IIf(IsWindow(g_TaskBarHwnd) = APIFALSE, False, True)
End Function


'--------------------------------------------------------------------------------
' Procedure  :       GetEdge
' Description:       Determines which the taskbar rect is on
' Parameters :       rc (RECT)
'--------------------------------------------------------------------------------
Private Function GetEdge(rc As RECT) As Long

    Dim uEdge As Long: uEdge = -1

    If (rc.Top = rc.Left) And (rc.bottom > rc.Right) Then
        uEdge = ABE_LEFT
    ElseIf (rc.Top = rc.Left) And (rc.bottom < rc.Right) Then
        uEdge = ABE_TOP
    ElseIf (rc.Top > rc.Left) Then
        uEdge = abe_bottom
    Else
        uEdge = ABE_RIGHT
    End If
    
    GetEdge = uEdge

End Function

