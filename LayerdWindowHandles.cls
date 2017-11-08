VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LayerdWindowHandles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'    Component  : LayerdWindowHandles
'    Project    : ViOrb5
'
'    Description: Contains all properties needed to transform a window into a
'                 layered window and back again
'
'    Modified   :
'--------------------------------------------------------------------------------
Option Explicit

Public theDC      As Long

Public oldBitmap  As Long

Public mainBitmap As Long

Private winSize   As win.SIZEL

Private srcPoint  As win.POINTL

Private released  As Boolean

Private m_myBlend As BLENDFUNCTION

Public Sub Release()

    If released Then Exit Sub
    released = True

    SelectObject theDC, oldBitmap
    DeleteObject mainBitmap
    DeleteObject oldBitmap
    DeleteDC theDC
End Sub

Public Function SetPoint(newPoint As win.POINTL)
    srcPoint = newPoint
End Function

Public Function SetSize(newSize As win.SIZEL)
    winSize = newSize
End Function

Public Function Update(ByVal hWnd As Long, _
                       ByVal hdc As Long, _
                       Optional alphaBlend As Byte = 255)
    m_myBlend.SourceConstantAlpha = alphaBlend
    Call UpdateLayeredWindow(hWnd, hdc, ByVal 0&, winSize, theDC, srcPoint, 0, m_myBlend, ULW_ALPHA)
End Function

Private Sub Class_Initialize()

    With m_myBlend
        .AlphaFormat = AC_SRC_ALPHA
        .BlendFlags = 0
        .BlendOp = AC_SRC_OVER
        .SourceConstantAlpha = 255
    End With
    
    released = False
End Sub

Private Sub Class_Terminate()
    Release
End Sub

