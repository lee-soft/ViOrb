Attribute VB_Name = "LayerdWindowSupport"
'--------------------------------------------------------------------------------
'    Component  : LayerdWindowSupport
'    Project    : ViOrb5
'
'    Description: Containts helper functions for layered windows
'
'    Modified   :
'--------------------------------------------------------------------------------
Option Explicit

Public Declare Function UpdateLayeredWindow _
               Lib "user32.dll" (ByVal hWnd As Long, _
                                 ByVal hdcDst As Long, _
                                 pptDst As Any, _
                                 psize As Any, _
                                 ByVal hdcSrc As Long, _
                                 pptSrc As Any, _
                                 ByVal crKey As Long, _
                                 ByRef pblend As BLENDFUNCTION, _
                                 ByVal dwFlags As Long) As Long

Private m_layeredAttrBank As Collection

Public Const ULW_ALPHA = &H2

Public Const WS_EX_LAYERED = &H80000

Public Const AC_SRC_ALPHA As Long = &H1

Public Const AC_SRC_OVER = &H0

Public Type BLENDFUNCTION

    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte

End Type

Public Enum AnchorPointConstants

    apTopLeft = 1
    apTop = 5
    apBottomLeft = 2
    apLeft = 6
    apBottomRight = 3
    apBottom = 7
    apTopRight = 4
    apRight = 8
    apMiddle = 9

End Enum

Public Function MakeLayerdWindow(ByRef sourceForm As Form, _
                                 Optional fromExistingLayeredWindow As Boolean = True, _
                                 Optional clickThrough As Boolean = False) As LayerdWindowHandles

    Dim KeyName As String

    KeyName = sourceForm.hWnd & "_hwnd"

    If m_layeredAttrBank Is Nothing Then
        Set m_layeredAttrBank = New Collection
    End If
    
    If ExistInCol(m_layeredAttrBank, KeyName) Then
        If fromExistingLayeredWindow Then
            m_layeredAttrBank(KeyName).Release
            m_layeredAttrBank.Remove KeyName
        Else
            Set MakeLayerdWindow = m_layeredAttrBank(KeyName)
            Call SetWindowLong(sourceForm.hWnd, GWL_EXSTYLE, GetWindowLong(sourceForm.hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED)

            Exit Function

        End If
    End If

    Dim srcPoint   As win.POINTL

    Dim winSize    As win.SIZEL

    Dim mDC        As Long

    Dim tempBI     As BITMAPINFO

    Dim mainBitmap As Long

    Dim oldBitmap  As Long

    Dim theHandles As New LayerdWindowHandles

    Dim newStyle   As Long

    m_layeredAttrBank.Add theHandles, sourceForm.hWnd & "_hwnd"

    With tempBI.bmiHeader
        .biSize = Len(tempBI.bmiHeader)
        .biBitCount = 32    ' Each pixel is 32 bit's wide
        .biHeight = sourceForm.ScaleHeight  ' Height of the form
        .biWidth = sourceForm.ScaleWidth    ' Width of the form
        .biPlanes = 1   ' Always set to 1
        .biSizeImage = .biWidth * .biHeight * (.biBitCount / 8) ' This is the number of bytes that the bitmap takes up. It is equal to the Width*Height*ByteCount (bitCount/8)
    End With
    
    mDC = CreateCompatibleDC(sourceForm.hdc)
    mainBitmap = CreateDIBSection(mDC, tempBI, DIB_RGB_COLORS, ByVal 0, 0, 0)
    
    If mainBitmap = 0 Then
        MsgBox "CreateDIBSection Failed", vbCritical

        Exit Function

    End If
    
    oldBitmap = SelectObject(mDC, mainBitmap)   ' Select the new bitmap, track the old that was selected
    
    If oldBitmap = 0 Then

        'MsgBox "SelectObject Failed", vbCritical
        Exit Function

    End If
    
    newStyle = GetWindowLong(sourceForm.hWnd, GWL_EXSTYLE)
    newStyle = newStyle Or WS_EX_LAYERED
    
    If (clickThrough) Then
        newStyle = newStyle Or WS_EX_TRANSPARENT
    End If
    
    If SetWindowLong(sourceForm.hWnd, GWL_EXSTYLE, newStyle) = 0 Then
        'MsgBox "Failed to create layered window!"
        'Exit Function
    End If
    
    ' Needed for updateLayeredWindow call
    srcPoint.x = 0
    srcPoint.y = 0
    winSize.cx = sourceForm.ScaleWidth
    winSize.cy = sourceForm.ScaleHeight
    
    theHandles.mainBitmap = mainBitmap
    theHandles.oldBitmap = oldBitmap
    theHandles.theDC = mDC
    
    theHandles.SetSize winSize
    theHandles.SetPoint srcPoint
    'theHandles.
    Set MakeLayerdWindow = theHandles

Handler:
End Function

