Attribute VB_Name = "FileSelecterHelper"
Option Explicit

'Source      :  http://visualbasic.happycodings.com/common-dialogs/code6.html
'Purpose     :  Allows the user to select a file name from a local or network directory.
'Inputs      :  sInitDir            The initial directory of the file dialog.
'               sFileFilters        A file filter string, with the following format:
'                                   eg. "Excel Files;*.xls|Text Files;*.txt|Word Files;*.doc"
'               [sTitle]            The dialog title
'               [lParentHwnd]       The handle to the parent dialog that is calling this function.
'Outputs     :  Returns the selected path and file name or a zero length string if the user pressed cancel
Function BrowseForFile(sInitDir As String, _
                       Optional ByVal sFileFilters As String, _
                       Optional sTitle As String = "Open File", _
                       Optional lParentHwnd As Long) As String
    
    Dim tFileBrowse As OPENFILENAME

    Const clMaxLen  As Long = 5000
    
    Dim theBuffer   As String

    theBuffer = String$(clMaxLen, Chr$(0))
    
    tFileBrowse.lStructSize = Len(tFileBrowse)
    
    'Replace friendly deliminators with nulls
    sFileFilters = Replace(sFileFilters, "|", vbNullChar)
    sFileFilters = Replace(sFileFilters, ";", vbNullChar)

    If Right$(sFileFilters, 1) <> vbNullChar Then
        'Add final delimiter
        sFileFilters = sFileFilters & vbNullChar
    End If
    
    'Select a filter
    tFileBrowse.lpstrFilter = StrPtr(sFileFilters & "All Files (*.*)" & vbNullChar & "*.*" & vbNullChar)
    'create a buffer for the file
    tFileBrowse.lpstrFile = StrPtr(theBuffer)
    'set the maximum length of a returned file
    tFileBrowse.nMaxFile = clMaxLen + 1
    'Create a buffer for the file title
    tFileBrowse.lpstrFileTitle = StrPtr(Space$(clMaxLen))
    'Set the maximum length of a returned file title
    tFileBrowse.nMaxFileTitle = clMaxLen + 1
    'Set the initial directory
    tFileBrowse.lpstrInitialDir = StrPtr(sInitDir)
    'Set the parent handle
    tFileBrowse.hwndOwner = lParentHwnd
    'Set the title
    tFileBrowse.lpstrTitle = StrPtr(sTitle)
    
    'No flags
    tFileBrowse.flags = 0

    'Show the dialog
    If GetOpenFileName(tFileBrowse) Then
        BrowseForFile = Trim$(GetString((tFileBrowse.lpstrFile)))

        If Right$(BrowseForFile, 1) = vbNullChar Then
            'Remove trailing null
            BrowseForFile = Left$(BrowseForFile, Len(BrowseForFile) - 1)
        End If
    End If

End Function

Public Function GetString(ByVal PtrStr As Long) As String

    Dim StrBuff As String * 256

    'Check for zero address
    If PtrStr = 0 Then
        GetString = vbNullString

        Exit Function

    End If

    'Copy data from PtrStr to buffer.
    win.CopyMemory ByVal StrBuff, ByVal PtrStr, 256
    'Strip any trailing nulls from string.
    GetString = StripNulls(StrBuff)
End Function

Public Function StripNulls(OriginalStr As String) As String

    'Strip any trailing nulls from input string.
    If (InStr(OriginalStr, Chr$(0)) > 0) Then
        OriginalStr = Left$(OriginalStr, InStr(OriginalStr, Chr$(0)) - 1)
    End If

    'Return modified string.
    StripNulls = OriginalStr
End Function


