VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ViOrb Options"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7260
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   378
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   484
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer timDelayRefresh 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4200
      Top             =   1560
   End
   Begin VB.Timer timDelayMoveMode 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4080
      Top             =   960
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   495
      Left            =   1440
      TabIndex        =   18
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse"
      Height          =   495
      Left            =   2640
      TabIndex        =   17
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton cmdMore 
      Caption         =   "&Get More"
      Height          =   495
      Left            =   240
      TabIndex        =   16
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CheckBox chkSpoof 
      Caption         =   "Spoof ViOrb as ViGlance Orb to maintain compatability with other ViApps"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   4800
      Width           =   6975
   End
   Begin VB.CheckBox chkManual 
      Caption         =   "Manually Position Orb"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   4320
      Width           =   2895
   End
   Begin VB.CommandButton cmdNudgeUp 
      Caption         =   "é"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   9
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   6
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton cmdNudgeBottom 
      Caption         =   "ê"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   9
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   5
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton cmdNudgeRight 
      Caption         =   "è"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   9
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   6480
      TabIndex        =   4
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton cmdNudgeLeft 
      Caption         =   "ç"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   9
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   4320
      TabIndex        =   3
      Top             =   1200
      Width           =   375
   End
   Begin VB.CheckBox chkStart 
      Caption         =   "Start With Windows"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   3720
      Width           =   2895
   End
   Begin VB.FileListBox flSkins 
      Appearance      =   0  'Flat
      Height          =   2340
      Left            =   240
      Pattern         =   "*.png"
      TabIndex        =   1
      Top             =   600
      Width           =   3495
   End
   Begin VB.Label Label3 
      Caption         =   "Don't adjust this setting unless you understand it."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   5160
      Width           =   6735
   End
   Begin VB.Label Label2 
      Caption         =   "Remember to makesure that 'manually position orb' is unchecked if you want to enable the nudge feature."
      Height          =   975
      Left            =   3960
      TabIndex        =   13
      Top             =   3600
      Width           =   3135
   End
   Begin VB.Label Label6 
      Caption         =   "Nudge Orb Tool"
      Height          =   375
      Left            =   4080
      TabIndex        =   11
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label lblTop 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   10
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lblBottom 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   9
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label lblRight 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   8
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label lblLeft 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   7
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Select start button"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'    Component  : frmOptions
'    Project    : ViOrb5
'
'    Description: Provides program options for ViOrb such as skin and startup
'                 options
'
'    Modified   :
'--------------------------------------------------------------------------------
Option Explicit

Private Const StartWithWindowsRegistryPath As String = "HKCU\Software\Microsoft\Windows\CurrentVersion\Run\"

Private m_ignoreInput                      As Boolean 'when true, events are effectively ignored when fired on controls

Sub ForceRefresh()
    Unload frmStartButton
    Unload frmZOrderKeeper
    Unload frmFader
    
    Set frmZOrderKeeper = Nothing
    Set frmStartButton = Nothing
    Set frmFader = Nothing

    frmZOrderKeeper.Show
    UpdateCaptions
End Sub

Sub SelectedSkinByName(ByVal theSkinName As String)

    On Error GoTo Handler

    If theSkinName = vbNullString Then

        Exit Sub

    End If
    
    Dim skinIndex As Long

    For skinIndex = 0 To flSkins.ListCount

        If UCase$(flSkins.List(skinIndex)) = UCase$(theSkinName) Then
            flSkins.Selected(skinIndex) = True
        End If

    Next

    Exit Sub

Handler:
    LogError 0, Err.Description, "frmOptions"
End Sub

Sub TestStartWithWindows()
    m_ignoreInput = True 'this is a very crude way to block events when programatically changing control properties
    chkStart.Value = vbUnchecked
    
    If UCase$(RegistryHelper.ReadKeyString(StartWithWindowsRegistryPath & App.Title)) = UCase$(App.Path & "\" & App.EXEName & ".exe") Then
        chkStart.Value = vbChecked
    End If

    m_ignoreInput = False
    
End Sub

Sub UpdateCaptions()
    lblLeft.Caption = frmStartButton.NudgeX
    lblRight.Caption = 0 - frmStartButton.NudgeX
    lblTop.Caption = frmStartButton.NudgeY
    lblBottom.Caption = 0 - frmStartButton.NudgeY
    
    m_ignoreInput = True 'this is a very crude way to block events when programatically changing control properties
    
    If frmStartButton.ManualPosition Then
        chkManual.Value = vbChecked
        
        cmdNudgeLeft.Enabled = False
        cmdNudgeRight.Enabled = False
        cmdNudgeUp.Enabled = False
        cmdNudgeBottom.Enabled = False
    Else
        chkManual.Value = vbUnchecked

        cmdNudgeLeft.Enabled = True
        cmdNudgeRight.Enabled = True
        cmdNudgeUp.Enabled = True
        cmdNudgeBottom.Enabled = True
    End If
    
    m_ignoreInput = False
End Sub

Sub cmdBrowse_Click()

    Dim selectedFile       As Scripting.File

    Dim szSelectedFile     As String

    Dim szProposedFileName As String

    Dim copyNumber         As Long

    szSelectedFile = BrowseForFile(0, "Portable Network Graphics;*.png", "Choose New Start Orb Image", Me.hWnd)

    If Not FSO.FileExists(szSelectedFile) Then

        Exit Sub

    End If
    
    Set selectedFile = FSO.GetFile(szSelectedFile)
    szProposedFileName = selectedFile.Name
    
    While FSO.FileExists(App.Path & "\Resources\" & szProposedFileName)

        copyNumber = copyNumber + 1
        szProposedFileName = FSO.GetBaseName(szSelectedFile) & "(" & copyNumber & ").png"

    Wend
    
    selectedFile.Copy App.Path & "\Resources\" & szProposedFileName, False
    flSkins.Refresh

    SelectedSkinByName szProposedFileName
End Sub

Private Sub chkManual_Click()

    If m_ignoreInput Then Exit Sub

    If chkManual.Value = vbUnchecked Then
        If IsWindow(TaskbarHelper.g_StartButtonHwnd) = APIFALSE Then
            If IsWindow(TaskbarHelper.g_viOrbToolbar) = APIFALSE Then
                frmInstall.Show vbModal
            End If

            If IsWindow(TaskbarHelper.g_viOrbToolbar) = APIFALSE Then
                chkManual.Value = vbChecked

                Exit Sub

            End If
            
            frmZOrderKeeper.HijackZOrder
        End If

        frmStartButton.ResetPosition
        
    Else

        If IsWindow(TaskbarHelper.g_viOrbToolbar) = APITRUE Then
            MsgBox "Remove the 'Start' Toolbar from your taskbar first." & vbCrLf & "Right click taskbar, Toolbars > 'Start'", vbCritical
            chkManual.Value = vbUnchecked

            Exit Sub

        End If
        
        frmStartButton.ActivateMoveMode
    End If

End Sub

Private Sub chkSpoof_Click()

    If m_ignoreInput Then Exit Sub

    If chkSpoof.Value = vbUnchecked Then
        frmStartButton.Caption = MainHelper.ViOrb_Identifier
    Else
        frmStartButton.Caption = MainHelper.ViGlance_Identifier
    End If

End Sub

Private Sub chkStart_Click()

    If m_ignoreInput Then Exit Sub

    If chkStart.Value = vbChecked Then
        RegistryHelper.WriteRegistryString StartWithWindowsRegistryPath & App.Title, App.Path & "\" & App.EXEName & ".exe"
    Else
        RegistryHelper.DeleteKey App.Path & "\" & App.EXEName & ".exe"
    End If

End Sub

Private Sub cmdDelete_Click()
    
    Dim szFilename As String

    szFilename = flSkins.List(flSkins.ListIndex)
    
    If LCase$(szFilename) = LCase$(Default_Orb_Name) Then

        Exit Sub

    End If
    
    SelectedSkinByName Default_Orb_Name

    If Not flSkins.List(flSkins.ListIndex) = Default_Orb_Name Then Exit Sub
    If Not FSO.FileExists(App.Path & "\Resources\" & szFilename) Then Exit Sub
    
    If MsgBox("Do you really want to delete, " & szFilename & "?", vbQuestion Or vbYesNo) = vbNo Then

        Exit Sub

    End If
    
    FSO.DeleteFile App.Path & "\Resources\" & szFilename, True
    flSkins.Refresh
    
    Me.Show
End Sub

Private Sub cmdMore_Click()
    Call ShellExecute(Me.hWnd, vbNullString, "http://lee-soft.com/skins/", vbNullString, App.Path, 0)
End Sub

Private Sub cmdNudgeBottom_Click()
    frmStartButton.NudgeY = frmStartButton.NudgeY + 1
    UpdateCaptions
End Sub

Private Sub cmdNudgeLeft_Click()
    frmStartButton.NudgeX = frmStartButton.NudgeX - 1
    UpdateCaptions
End Sub

Private Sub cmdNudgeRight_Click()
    frmStartButton.NudgeX = frmStartButton.NudgeX + 1
    UpdateCaptions
End Sub

Private Sub cmdNudgeUp_Click()
    frmStartButton.NudgeY = frmStartButton.NudgeY - 1
    UpdateCaptions
End Sub

Private Sub flSkins_Click()

    If m_ignoreInput Then Exit Sub
    
    RegistryHelper.WriteRegistryString AppSettingsRegistryPath & "filename", flSkins.List(flSkins.ListIndex)
    ForceRefresh
End Sub

Private Sub Form_Activate()
    UpdateCaptions
End Sub

Private Sub Form_Load()
    flSkins.Path = App.Path & "\Resources\"
    TestStartWithWindows
    
    If frmStartButton.Caption = MainHelper.ViGlance_Identifier Then
        chkSpoof.Value = vbChecked
    End If
    
    UpdateCaptions
End Sub

Private Sub timDelayMoveMode_Timer()
    timDelayMoveMode.Enabled = False
    frmStartButton.ActivateMoveMode
End Sub

Private Sub timDelayRefresh_Timer()
        
    timDelayRefresh.Enabled = False
    ForceRefresh
End Sub
