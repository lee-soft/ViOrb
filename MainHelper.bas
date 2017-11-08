Attribute VB_Name = "MainHelper"
'--------------------------------------------------------------------------------
'    Component  : MainHelper
'    Project    : ViOrb5
'
'    Description: Program entry point container
'
'    Modified   :
'--------------------------------------------------------------------------------
Option Explicit

Public Const AppSettingsRegistryPath As String = "HKCU\Software\ViOrb\"

Public Const ViGlance_Identifier     As String = "#Start~ViGlance#"

Public Const ViOrb_Identifier        As String = "#Start~ViOrb#"

Public Const Default_Orb_Name        As String = "Windows 7 Orb.png"

Private m_showOptions                As Boolean


'--------------------------------------------------------------------------------
' Procedure  :       Main
' Description:       Program entry point
' Parameters :
'--------------------------------------------------------------------------------
Sub Main()

    InitializeGDIIfNotInitialized
    DetermineWindowsVersion_IfNeeded
    
    If Not WaitForTaskbar Then
        If MsgBox("Unable to locate the Windows taskbar. Do you really want to continue?", vbYesNo Or vbQuestion) = vbNo Then

            Exit Sub

        End If
    End If
    
    'first run is implied with no saved registry, so show options
    If RegistryHelper.ReadKeyString(AppSettingsRegistryPath & "filename") = vbNullString Then
        m_showOptions = True
    End If
    
    TaskbarHelper.UpdatehWnds
    
    If IsWindow(TaskbarHelper.g_StartButtonHwnd) = APIFALSE Then
        If IsWindow(TaskbarHelper.g_viOrbToolbar) = APIFALSE And ReadKeyString(AppSettingsRegistryPath & "noorb_warning") <> "1" Then
            frmInstall.Show vbModal
        End If
    End If

    If RegistryHelper.ReadKeyInteger(AppSettingsRegistryPath & "no_splash", 0) = 0 Then

        Dim theSplashScreen As New frmSplash

        theSplashScreen.Show vbModal
    End If
    
    Load frmClock
    frmZOrderKeeper.Show
    
    If m_showOptions Then
        frmOptions.Show
    End If

End Sub

