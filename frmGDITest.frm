VERSION 5.00
Begin VB.Form frmGDITest 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   480
   ClientWidth     =   4530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   193
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   302
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   3360
      TabIndex        =   0
      Top             =   2040
      Width           =   1095
   End
End
Attribute VB_Name = "frmGDITest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Lee Chantrey
'Description: Simple GDI layered window test

Option Explicit

Private m_gfx           As New GDIPGraphics

Private m_image         As New GDIPImage

Private m_layeredWindow As LayerdWindowHandles

Private Sub Command1_Click()
    Set m_layeredWindow = MakeLayerdWindow(Me)
    m_gfx.FromHDC m_layeredWindow.theDC
    
    m_gfx.Clear
    m_gfx.DrawImageRectFv m_image, 0, 0, m_image.Width, m_image.Height
    'Me.Refresh
    m_layeredWindow.Update Me.hWnd, Me.hdc, 255
    
    StayOnTop Me, True
End Sub

Private Sub Form_Load()
    GDIPlusCreate
    m_image.FromFile "D:\Users\littlelee\Documents\Projects\ViOrb 5\Resources\Windows 7 Orb.png"
    
End Sub
