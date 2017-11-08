Attribute VB_Name = "GeneralHelper"
'--------------------------------------------------------------------------------
'    Component  : GeneralHelper
'    Project    : ViOrb5
'
'    Description: If a function or API decleration needs a place to live and it
'                 doesn't seem to belong anywhere else it will go here.
'
'                 TODO: Break up this module into slimmer dedicated modules like
'                 RectHelper, even if it only contains 1 declaration/routine
'
'
'    Modified   :
'--------------------------------------------------------------------------------
Option Explicit

Public Declare Function GetTopWindow Lib "user32.dll" (ByVal hWnd As Long) As Long

Public Declare Function GetNextWindow _
               Lib "user32.dll" _
               Alias "GetWindow" (ByVal hWnd As Long, _
                                  ByVal wFlag As Long) As Long

Public Declare Function SHAppBarMessage _
               Lib "shell32.dll" (ByVal dwMessage As Long, _
                                  ByRef pData As APPBARDATA) As Long

Public Declare Sub CopyMemory _
               Lib "kernel32.dll" _
               Alias "RtlMoveMemory" (ByRef Destination As Any, _
                                      ByRef Source As Any, _
                                      ByVal Length As Long)

Public Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long

Private Declare Function SendInput _
                Lib "user32.dll" (ByVal nInputs As Long, _
                                  pInputs As GENERALINPUT, _
                                  ByVal cbSize As Long) As Long

Public Declare Function TrackMouseEvent _
               Lib "user32" (lpEventTrack As TrackMouseEvent) As Long

Private Type KEYBDINPUT

    wVk As Integer
    wScan As Integer
    dwFlags As Long
    time As Long
    dwExtraInfo As Long

End Type

Private Type SHELLEXECUTEINFOW

    cbSize As Long
    fMask As Long
    hWnd As Long
    lpVerb As Long
    lpFile As Long
    lpParameters As Long
    lpDirectory As Long
    nShow As Long
    hInstApp As Long
    ' fields
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    HICON As Long
    hProcess As Long

End Type

Private Type HARDWAREINPUT

    uMsg As Long
    wParamL As Integer
    wParamH As Integer

End Type

Private Type GENERALINPUT

    dwType As Long
    xi(0 To 23) As Byte

End Type

' Can be used with either W or A functions
' Pass VarPtr(wfd) to W or simply wfd to A
Private Type WIN32_FIND_DATAW

    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * 260
    cAlternate As String * 14

End Type

Public Const RDW_ALLCHILDREN = &H80

Public Const RDW_ERASE = &H4

Public Const RDW_INVALIDATE = &H1

Public Const RDW_UPDATENOW = &H100

Public Const WM_SYSMENU        As Long = &H313

Public Const ABM_GETTASKBARPOS As Long = &H5

Public Enum AbeBarEnum

    abe_bottom = 3
    ABE_LEFT = 0
    ABE_RIGHT = 2
    ABE_TOP = 1

End Enum

Public Const AW_CENTER       As Long = &H10

Public Const AW_SLIDE        As Long = &H40000

Public Const AW_HIDE         As Long = &H10000

Public Const AW_BLEND        As Long = &H80000

Public Const AW_VER_NEGATIVE As Long = &H8

Public Const AW_VER_POSITIVE As Long = &H4

Public Const HSHELL_REDRAW   As Long = 6

Public Const HSHELL_HIGHBIT = &H8000

Public Const HSHELL_FLASH = 32774

Public Const HSHELL_WINDOWDESTROYED As Long = 2

Public Const HSHELL_WINDOWCREATED   As Long = 1

Public Const HSHELL_WINDOWACTIVATED As Long = 4

Public Const IDANI_OPEN = &H1

Public Const IDANI_CLOSE = &H2

Public Const IDANI_CAPTION = &H3

Public Const ULW_OPAQUE = &H4

Public Const ULW_COLORKEY = &H1

Public Const GCL_HICON = (-14)

Public Const GCL_HICONSM = (-34)

Public Const TME_LEAVE        As Long = &H2

Public Const WS_EX_NOACTIVATE As Long = &H8000000

Public Const SMTO_BLOCK = &H1

Public Const SMTO_ABORTIFHUNG = &H2

Private Const KEYEVENTF_KEYUP = &H2

Private Const INPUT_KEYBOARD = 1

Private Const INPUT_HARDWARE = 2

' API Defined Types
Public Type APPBARDATA

    cbSize As Long
    hWnd As Long
    uCallbackMessage As Long
    uEdge As Long
    rc As RECT
    lParam As Long

End Type

Public Enum PNG_DRAWMODE

    NOT_SET = 0
    LAYERED_MODE = 1
    NORMAL_MODE = 2

End Enum

Private m_GDIInitialized As Boolean

Private m_FSO            As FileSystemObject

Public g_WindowsVersion  As OSVERSIONINFO

Public g_WindowsXP       As Boolean

'Public g_WindowsVista    As Boolean #doesn't matter

'Public g_Windows7        As Boolean #doesn't matter

Public g_Windows8        As Boolean

Public g_Windows81       As Boolean

Public ViOrbToolbar      As Boolean

Function DetermineWindowsVersion_IfNeeded()

    Dim winRegistryVersion As String

    If g_WindowsVersion.dwBuildNumber <> 0 Then

        Exit Function

    End If

    g_WindowsVersion = GetWindowsOSVersion()

    g_WindowsXP = False
    'g_WindowsVista = False
    'g_Windows7 = False
    g_Windows8 = False
    g_Windows81 = False
    
    winRegistryVersion = RegistryHelper.ReadKeyString("HKLM\Software\Microsoft\Windows NT\CurrentVersion\CurrentVersion")
    
    If g_WindowsVersion.dwMajorVersion = 5 Then
        If g_WindowsVersion.dwMinorVersion = 1 Or g_WindowsVersion.dwMinorVersion = 2 Then
            g_WindowsXP = True
        End If

    ElseIf g_WindowsVersion.dwMajorVersion = 6 Then

        If g_WindowsVersion.dwMinorVersion = 0 Then
            'g_WindowsVista = True //nobody cares
        ElseIf g_WindowsVersion.dwMinorVersion = 1 Then
            'g_Windows7 = True //nobody cares
        ElseIf g_WindowsVersion.dwMinorVersion = 2 Then
            'Determine Windows 8 Version
            g_Windows8 = True
            
            If winRegistryVersion = "6.2" Then
                
            ElseIf winRegistryVersion = "6.3" Then
                g_Windows81 = True
            Else
                MsgBox "This version of Windows is unknown.. " & App.Title & " may not behave as expected!", vbCritical
                g_Windows8 = True
            End If

        Else
            MsgBox "This version of Windows is unknown.. " & App.Title & " may not behave as expected!", vbCritical
            g_Windows8 = True
        End If

    Else
        MsgBox "This version of Windows is unknown.. " & App.Title & " may not behave as expected!", vbCritical
        g_Windows8 = True
    End If
    
End Function


'--------------------------------------------------------------------------------
' Procedure  :       DisposeGDIIfLast
' Description:       Everytime a form is unloaded it should call this function
'                    then it ensures the GDI is properly disposed.
' Parameters :
'--------------------------------------------------------------------------------
Public Function DisposeGDIIfLast()

    If Forms.count = 1 Then
        GDIPlusDispose
    End If

End Function

'--------------------------------------------------------------------------------
' Procedure  :       EnumTaskbarChildrenToFindStartButton
' Description:       Determines if the window handle is the windows start button
'                    and if it isn't it continues to enumerate the windows
' Parameters :       lhWnd (Long)
'                    lParam (Long)
'--------------------------------------------------------------------------------
Function EnumTaskbarChildrenToFindStartButton(ByVal lhWnd As Long, _
                                              ByVal lParam As Long) As Long
       

    Dim WinClassBuf As String * 255, WinTitleBuf As String * 255
    Dim WinClass    As String, WinTitle As String

    WinClass = StripNulls(WinClassBuf)  ' remove extra Nulls & spaces
    WinTitle = StripNulls(WinTitleBuf)
 
    If LCase$(WinClass) = "start" And LCase$(WinTitle) = "start" Then
        lParam = lhWnd
        g_StartButtonHwnd = lhWnd
            
        EnumTaskbarChildrenToFindStartButton = False
    Else
        EnumTaskbarChildrenToFindStartButton = True
    End If

End Function

'--------------------------------------------------------------------------------
' Procedure  :       ExistInCol
' Description:       Tests if there's an element at the given key/index

' Parameters :       cTarget (Collection)
'                    sKey (Variant)
'--------------------------------------------------------------------------------
Public Function ExistInCol(ByRef cTarget As Collection, sKey) As Boolean

    On Error GoTo Handler

    ExistInCol = Not (IsEmpty(cTarget(sKey)))
    
    Exit Function

Handler:
    ExistInCol = False
End Function


'--------------------------------------------------------------------------------
' Procedure  :       FileExists
' Description:       Tests if the given path is a valid file
' Parameters :       fileName (String)
'--------------------------------------------------------------------------------
Function FileExists(fileName As String) As Boolean

    On Error GoTo ErrorHandler

    ' get the attributes and ensure that it isn't a directory
    FileExists = (GetAttr(fileName) And vbDirectory) = 0
ErrorHandler:
    ' if an error occurs, this function returns False
End Function


'--------------------------------------------------------------------------------
' Procedure  :       FSO
' Description:       A method that always promises to return a file system
'                    object
'
' Parameters :
'--------------------------------------------------------------------------------
Public Function FSO() As FileSystemObject

    If m_FSO Is Nothing Then
        Set m_FSO = New FileSystemObject
    End If
    
    Set FSO = m_FSO
End Function

Public Function GetCursorPoint() As win.POINTL

    Dim thisPos As win.POINTL
    
    GetCursorPos thisPos
    GetCursorPoint = thisPos
End Function


'--------------------------------------------------------------------------------
' Procedure  :       GetFilenameFromPath
' Description:       Regardless of the validity of the path, it returns the
'                    the filename part (everything past the final last \)
' Parameters :       FullPath (String)
'--------------------------------------------------------------------------------
Public Function GetFilenameFromPath(FullPath As String) As String

    GetFilenameFromPath = Right$(FullPath, Len(FullPath) - InStrRev(FullPath, "\"))
End Function


'--------------------------------------------------------------------------------
' Procedure  :       GetZOrder
' Description:       Attempts to retrieve the overlapping order of the target
'                    hWnd
' Parameters :       hWndTarget (Long)
'--------------------------------------------------------------------------------
Public Function GetZOrder(ByVal hWndTarget As Long) As Long
    
    Dim hWnd      As Long

    Dim lngZOrder As Long

    ' Loop through window list and
    ' compare to hWnd to hwndTarget to find global ZOrder
    hWnd = GetTopWindow(0)
    lngZOrder = 0
    
    Do While hWnd And hWnd <> hWndTarget
        ' Get next window and move on.
        hWnd = GetNextWindow(hWnd, GW_HWNDNEXT)
        lngZOrder = lngZOrder + 1
    Loop
    
    GetZOrder = lngZOrder

End Function


'--------------------------------------------------------------------------------
' Procedure  :       hWndBelongToUs
' Description:       A quick check to see if the given Window Handle is one of
'                    the forms in this application
'
' Parameters :       hWnd (Long)
'                    ExceptionHwnd (Long)
'--------------------------------------------------------------------------------
Public Function hWndBelongToUs(hWnd As Long, Optional ExceptionHwnd As Long) As Boolean

    Dim thisForm As Form

    hWndBelongToUs = False

    For Each thisForm In Forms

        If thisForm.hWnd = hWnd Then
            If hWnd = ExceptionHwnd Then
                hWndBelongToUs = False
            Else
                hWndBelongToUs = True
            End If
            
            Exit For

        End If

    Next
    
End Function

Public Function InitializeGDIIfNotInitialized() As Boolean

    If Not m_GDIInitialized Then

        ' Must call this before using any GDI+ call:
        If Not (GDIPlusCreate()) Then

            Exit Function

        End If
    
        m_GDIInitialized = True
    End If
    
    InitializeGDIIfNotInitialized = m_GDIInitialized
End Function


'--------------------------------------------------------------------------------
' Procedure  :       IsStyle
' Description:       Tests if a specific STYLE is applied
' Parameters :       lAll (Long) all the styles
'                    lBit (Long) style to test
'--------------------------------------------------------------------------------
Public Function IsStyle(ByVal lAll As Long, ByVal lBit As Long) As Boolean
      
    IsStyle = False

    If (lAll And lBit) = lBit Then
        IsStyle = True
    End If

End Function

Public Sub LogError(ByVal lNum As Long, ByVal sDesc As String, ByVal sFrom As String)
    
    Debug.Print "APP ERROR; " & sDesc & " ; " & sFrom
    
    Dim FileNum As Integer

    FileNum = FreeFile
    Open App.Path & "\errors.log" For Append As FileNum
    Write #FileNum, lNum, sDesc, sFrom, Now()
    Close FileNum
End Sub


'--------------------------------------------------------------------------------
' Procedure  :       SetKeyDown
' Description:       Simulates Key pressed down
' Parameters :       KeyCode (Long)
'--------------------------------------------------------------------------------
Public Function SetKeyDown(KeyCode As Long)

    Dim GInput(0 To 1) As GENERALINPUT

    Dim KInput         As KEYBDINPUT

    KInput.wVk = KeyCode 'the key we're going to press
    KInput.dwFlags = 0 'press the key
    'copy the structure into the input array's buffer.
    GInput(0).dwType = INPUT_KEYBOARD ' keyboard input
    CopyMemory GInput(0).xi(0), KInput, Len(KInput)
    
    'send the input now
    Call SendInput(2, GInput(0), Len(GInput(0)))

End Function

'--------------------------------------------------------------------------------
' Procedure  :       SetKeyUp
' Description:       Simulates key being released
' Parameters :       KeyCode (Long)
'--------------------------------------------------------------------------------
Public Function SetKeyUp(KeyCode As Long)

    Dim GInput(0 To 1) As GENERALINPUT

    Dim KInput         As KEYBDINPUT

    'do the same as above, but for releasing the key
    KInput.wVk = KeyCode ' the key we're going to realease
    KInput.dwFlags = KEYEVENTF_KEYUP ' release the key
    GInput(1).dwType = INPUT_KEYBOARD ' keyboard input
    CopyMemory GInput(1).xi(0), KInput, Len(KInput)
    'send the input now
    Call SendInput(2, GInput(0), Len(GInput(0)))

End Function


'--------------------------------------------------------------------------------
' Procedure  :       SetOwner
' Description:       [type_description_here]
' Parameters :       HwndtoUse (Variant)
'                    HwndofOwner (Variant)
'--------------------------------------------------------------------------------
Function SetOwner(ByVal HwndtoUse, ByVal HwndofOwner) As Long
    SetOwner = SetWindowLong(HwndtoUse, GWL_HWNDPARENT, HwndofOwner)
End Function

Public Sub StayOnTop(frmForm As Form, fOnTop As Boolean)

    Dim lState As Long

    Dim iLeft  As Integer, iTop As Integer, iWidth As Integer, iHeight As Integer

    With frmForm
        iLeft = .Left / Screen.TwipsPerPixelX
        iTop = .Top / Screen.TwipsPerPixelY
        iWidth = .Width / Screen.TwipsPerPixelX
        iHeight = .Height / Screen.TwipsPerPixelY
    End With

    If fOnTop Then
        lState = HWND_TOPMOST
    Else
        lState = HWND_NOTOPMOST
    End If

    'couldn't we use SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    'instead of doing the above sh*t
    'TODO: investigate and implement the above suggestion
    Call SetWindowPos(frmForm.hWnd, lState, iLeft, iTop, iWidth, iHeight, 0)
End Sub


'--------------------------------------------------------------------------------
' Procedure  :       UnloadApplication
' Description:       Safely unloads all the parent objects before ending
'                    the process
' Parameters :
'--------------------------------------------------------------------------------
Public Function UnloadApplication()

    Dim F As Form

    For Each F In Forms

        Unload F
    Next

    End

End Function

Private Function GetWindowsOSVersion() As OSVERSIONINFO

    Dim osv As OSVERSIONINFO

    osv.dwOSVersionInfoSize = Len(osv)
    
    If GetVersionEx(osv) = 1 Then
        GetWindowsOSVersion = osv
    End If

End Function


'--------------------------------------------------------------------------------
' Procedure  :       ShellCommand
' Description:       A crude way to launch a process not attached to your own
'                    process and simulates user opening the application
' Parameters :       Program (String)
'--------------------------------------------------------------------------------
Public Function ShellCommand(Program As String) As Boolean

    On Error GoTo Handler

    Shell "cmd.exe /c " & Program & " && exit", vbHide
    ShellCommand = True
    
    Exit Function

Handler:
    ShellCommand = False
End Function


