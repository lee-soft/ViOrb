VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBinaryFileStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'vbAccelerator - Contents of code file: cBinaryFileStream.cls

Option Explicit

Private m_sFile   As String

Private m_iFile   As Integer

Private m_iLen    As Long

Private m_iOffset As Long

Private Declare Sub CopyMemory _
                Lib "kernel32" _
                Alias "RtlMoveMemory" (lpvDest As Any, _
                                       lpvSource As Any, _
                                       ByVal cbCopy As Long)

Public Sub Dispose()

    If (m_iFile) Then
        Close #m_iFile
        m_iFile = 0
    End If

End Sub

Public Function Read(buffer() As Byte, ByVal readSize As Long) As Long
   
    Dim lReadSize As Long

    lReadSize = readSize

    If (m_iOffset + lReadSize >= m_iLen) Then
        readSize = m_iLen - m_iOffset

        If (readSize > 0) Then
            ReDim newBuffer(0 To readSize - 1) As Byte
            Get #m_iFile, , newBuffer
            CopyMemory buffer(0), newBuffer(0), readSize
        Else
            Dispose
        End If

        m_iOffset = m_iOffset + readSize
    Else
        ' Can read
        Get #m_iFile, , buffer
        m_iOffset = m_iOffset + readSize
    End If

    Read = readSize
   
End Function

Private Function FileExists(ByVal sFile As String, ByRef lErr As Long) As Boolean
   
    lErr = 0

    On Error Resume Next

    Dim sDir As String

    sDir = Dir$(sFile)
    lErr = Err.Number

    On Error GoTo 0
   
    If (lErr = 0) Then
        If (Len(sDir) > 0) Then
            FileExists = True
        Else
            lErr = 53
        End If
    End If
   
End Function

Private Sub Class_Terminate()
    Dispose
End Sub

Public Property Get File() As String
    File = m_sFile
End Property

Public Property Let File(ByVal sFile As String)
    Dispose
    m_sFile = sFile

    Dim lErr As Long

    If (FileExists(m_sFile, lErr)) Then
        m_iFile = FreeFile
        Open m_sFile For Binary Access Read Lock Write As #m_iFile
        m_iLen = LOF(m_iFile)
    Else
        Err.Raise lErr, App.EXEName & ".File"
    End If

End Property

Public Property Get Length() As Long
    Length = m_iLen
End Property
