Attribute VB_Name = "RectHelper"
'--------------------------------------------------------------------------------
'    Component  : RectHelper
'    Project    : ViOrb5
'
'    Description: Contains helper functions involding RECTs.
'                 (stripped for ViOrb)
'
'    Modified   :
'--------------------------------------------------------------------------------
Option Explicit


'--------------------------------------------------------------------------------
' Procedure  :       PointInsideOfRect
' Description:       Determines if the given point is indeed inside the given
'                    RECT
' Parameters :       srcPoint (win.POINTL)
'                    srcRect (win.RECT)
'--------------------------------------------------------------------------------
Public Function PointInsideOfRect(srcPoint As win.POINTL, srcRect As win.RECT) As Boolean

    PointInsideOfRect = False

    If srcPoint.y > srcRect.Top And srcPoint.y < srcRect.bottom And srcPoint.x > srcRect.Left And srcPoint.x < srcRect.Right Then

        PointInsideOfRect = True
    End If

End Function

