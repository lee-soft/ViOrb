Attribute VB_Name = "RegistryHelper"
'--------------------------------------------------------------------------------
'    Component  : RegistryHelper
'    Project    : ViOrb5
'
'    Description: A lazy way to read and write registry changes (via the shell
'                 the shell object)
'                 TODO: Import and use the proper API Wrapper Registry wrapper
'                 class and ditch this disgusting shell object wrapper
'
'    Modified   :
'--------------------------------------------------------------------------------
Option Explicit

Private m_shellClass As Object

Public Sub DeleteKey(Value As String)

    On Error Resume Next

    RegClass.RegDelete Value
End Sub

Public Function ReadKeyInteger(Value As String, default As Long) As Long

    On Error GoTo Handler

    ReadKeyInteger = CLng(RegClass.RegRead(Value))

    Exit Function

Handler:
    ReadKeyInteger = default
End Function

Public Function ReadKeyString(Value As String) As String

    On Error Resume Next

    ReadKeyString = CStr(RegClass.RegRead(Value))
End Function

Public Sub WriteRegistryInteger(Folder As String, Value As Long)

    On Error Resume Next

    RegClass.RegWrite Folder, Value, "REG_DWORD"
End Sub

Public Sub WriteRegistryString(Folder As String, Value As String)

    On Error Resume Next

    RegClass.RegWrite Folder, Value
End Sub

Private Function RegClass() As Object

    If m_shellClass Is Nothing Then
        Set m_shellClass = CreateObject("wscript.shell")
    End If
    
    Set RegClass = m_shellClass
End Function

