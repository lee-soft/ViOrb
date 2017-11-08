Attribute VB_Name = "ProcessHelper"
'--------------------------------------------------------------------------------
'    Component  : ProcessHelper
'    Project    : ViOrb5
'
'    Description: Contains process helper and management functions
'
'    Modified   :
'--------------------------------------------------------------------------------
Option Explicit


'--------------------------------------------------------------------------------
' Procedure  :       KillProcess
' Description:       A crude way to terminate all processes by image name
' Parameters :       theProcessName (String)
'--------------------------------------------------------------------------------
Public Function KillProcess(ByVal theProcessName As String)

    On Error Resume Next

    Dim objWMIService, objProcess, colProcess

    Dim strComputer, strProcessKill

    strComputer = "."
    strProcessKill = "'" & theProcessName & ".exe" & "'"
    
    Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
    
    Set colProcess = objWMIService.ExecQuery("Select * from Win32_Process Where Name = " & strProcessKill)

    For Each objProcess In colProcess

        objProcess.Terminate
    Next

    ' End of WMI Example of a Kill Process
End Function

'--------------------------------------------------------------------------------
' Procedure  :       ProcessCount
' Description:       Counts the number of instances of a given process name
' Parameters :       theProcessName (String)
'--------------------------------------------------------------------------------
Public Function ProcessCount(ByVal theProcessName As String) As Long

    On Error Resume Next

    Dim objWMIService, objProcess, colProcess

    Dim strComputer, strProcessKill

    strComputer = "."
    strProcessKill = "'" & theProcessName & ".exe" & "'"
    
    Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
    
    Set colProcess = objWMIService.ExecQuery("Select * from Win32_Process Where Name = " & strProcessKill)

    For Each objProcess In colProcess

        ProcessCount = ProcessCount + 1
    Next

    ' End of WMI Example of a Kill Process
End Function

