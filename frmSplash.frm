VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   4980
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   5250
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer timClose 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   300
      Top             =   2000
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Portions of this code are based on ViOrb 2 and ViGlance and ViPad."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   8
      Top             =   3240
      Width           =   4815
   End
   Begin VB.Image Image2 
      Height          =   720
      Left            =   240
      Picture         =   "frmSplash.frx":000C
      Top             =   120
      Width           =   675
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ViOrb Open"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   1080
      TabIndex        =   7
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Lee Matthew Chantrey of Lee-Soft.com is no way associated with Microsoft or Windows."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   6
      Top             =   1200
      Width           =   4815
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmSplash.frx":19CE
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      TabIndex        =   5
      Top             =   1800
      Width           =   4815
   End
   Begin VB.Label lblBottom 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "I hope you enjoy using ViOrb"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1200
      TabIndex        =   4
      Top             =   3960
      Width           =   2895
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Build 2988"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   360
      Width           =   3135
   End
   Begin VB.Label lblLink 
      BackColor       =   &H00FFFFFF&
      Caption         =   "www.lee-soft.com"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1920
      MouseIcon       =   "frmSplash.frx":1AF7
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label lblEmail 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "support@lee-soft.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1470
      TabIndex        =   1
      Top             =   4680
      Width           =   2325
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Lee-Soft.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   0
      Top             =   600
      UseMnemonic     =   0   'False
      Width           =   3615
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'    Component  : frmSplash
'    Project    : ViOrb5
'
'    Description: The boring and annoying splash screen that helped bring much
'                 needed revenue
'
'    Modified   :
'--------------------------------------------------------------------------------
Option Explicit

Private Sub Form_Click()

    If timClose.Enabled = False Then
        Unload Me
    End If

End Sub

Private Sub Form_Load()

    If Forms.count = 1 Then
        timClose.Enabled = True
    End If
    
    StayOnTop Me, True

    Label3.Caption = "Lee-Soft"
    Label2.Caption = "Version 5.0" & " (build " & App.Revision & ")"
End Sub

Private Sub Label1_Click()
    Call ShellExecute(Me.hWnd, vbNullString, "http://www.lee-soft.com/", vbNullString, App.Path, 0)
    Unload Me
End Sub

Private Sub Label2_Click()
    Call ShellExecute(Me.hWnd, vbNullString, "http://www.lee-soft.com/", vbNullString, App.Path, 0)
    Unload Me
End Sub

Private Sub Label3_Click()
    Call ShellExecute(Me.hWnd, vbNullString, "http://www.lee-soft.com/", vbNullString, App.Path, 0)
    Unload Me
End Sub

Private Sub Label4_Click()
    Unload Me
End Sub

Private Sub Label5_Click()
    Unload Me
End Sub

Private Sub Label6_Click()
    Unload Me
End Sub

Private Sub lblBottom_Click()
    Unload Me
End Sub

Private Sub lblLink_Click()
    Call ShellExecute(Me.hWnd, vbNullString, "http://www.lee-soft.com/", vbNullString, App.Path, 0)
    Unload Me
End Sub

Private Sub timClose_Timer()
    Unload Me
End Sub
