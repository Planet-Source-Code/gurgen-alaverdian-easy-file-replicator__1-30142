Attribute VB_Name = "fileState"
Public Container    As IADsContainer
Public NewLogFile   As String
Public FileToOpen   As String
Public noSound      As Boolean

Public Const SND_ASYNC = &H1
Public Const SND_MEMORY = &H4
Public Const SND_NODEFAULT = &H2
Public Const INVALID_HANDLE_VALUE = -1
Public Const MAX_PATH = 260
Public Const SHGFI_DISPLAYNAME = &H200

Public Type SHFILEINFO
   hIcon As Long
   iIcon As Long
   dwAttributes As Long
   szDisplayName As String * MAX_PATH
   szTypeName As String * 80
End Type

Public Type FILETIME
    dwLowDateTime   As Long
    dwHighDateTime  As Long
End Type

Public Type WIN32_FIND_DATA
   dwFileAttributes As Long
   ftCreationTime   As FILETIME
   ftLastAccessTime As FILETIME
   ftLastWriteTime  As FILETIME
   nFileSizeHigh    As Long
   nFileSizeLow     As Long
   dwReserved0      As Long
   dwReserved1      As Long
   cFileName        As String * MAX_PATH
   cAlternate       As String * 14
End Type

Public Type SECURITY_ATTRIBUTES
    nLength             As Long
    pSecurityDescriptor As Long
    bInheritHandle      As Long
End Type

Public Declare Function SHGetFileInfo Lib "shell32" _
                        Alias "SHGetFileInfoA" _
                        (ByVal pszPath As Any, _
                        ByVal dwFileAttributes As Long, _
                        psfi As SHFILEINFO, _
                        ByVal cbFileInfo As Long, _
                        ByVal uFlags As Long) As Long

Public Declare Function CreateDirectory Lib "kernel32" _
                        Alias "CreateDirectoryA" _
                        (ByVal lpPathName As String, _
                        lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
   
Public Declare Function CopyFile Lib "kernel32" _
                        Alias "CopyFileA" _
                        (ByVal lpExistingFileName As String, _
                        ByVal lpNewFileName As String, _
                        ByVal bFailIfExists As Long) As Long

Public Declare Function sndPlaySound Lib "winmm.dll" _
                        Alias "sndPlaySoundA" _
                        (lpszSoundName As Any, _
                        ByVal uFlags As Long) As Long

Public Declare Function FindFirstFile Lib "kernel32" _
                        Alias "FindFirstFileA" _
                        (ByVal lpFileName As String, _
                        lpFindFileData As WIN32_FIND_DATA) As Long
   
Public Declare Function FindClose Lib "kernel32" _
                        (ByVal hFindFile As Long) As Long
  
Public Declare Function FindNextFile Lib "kernel32" _
                         Alias "FindNextFileA" _
                        (ByVal hFindFile As Long, _
                        lpFindFileData As WIN32_FIND_DATA) As Long

Public Function FileExists(sFile As String) As Boolean
    Dim shfi As SHFILEINFO
    If SHGetFileInfo(ByVal sFile, 0&, shfi, Len(shfi), SHGFI_DISPLAYNAME) Then
        FileExists = True
        Else
        FileExists = False
    End If
End Function

'Public Function FileExists(sFile As String) As Boolean

'Dim FindData As WIN32_FIND_DATA
'Dim getFile As Long
   
'getFile = FindFirstFile(sFile, FindData)
'FileExists = getFile <> INVALID_HANDLE_VALUE
'Call FindClose(getFile)
'End Function

Public Function ReplicateFiles(absPath As String, _
                             absDest As String, _
                             Override As Boolean) As Long

Dim FindData As WIN32_FIND_DATA
Dim SA          As SECURITY_ATTRIBUTES
Dim getFile     As Long
Dim nextFile    As Long
Dim firstFile   As String
Dim fPath       As String

fPath = Left$(absPath, InStrRev(absPath, "\"))
getFile = FindFirstFile(absPath, FindData)
If getFile Then
    Do
        firstFile = Left$(FindData.cFileName, InStr(FindData.cFileName, Chr$(0)))
        Call CopyFile(fPath & firstFile, absDest & firstFile, Override)
        nextFile = FindNextFile(getFile, FindData)
    Loop Until nextFile = 0
End If
Call FindClose(getFile)
End Function

Public Function OpenFile(FileToOpen As String)

Dim s           As Integer
Dim I           As Integer
Dim J           As Integer
Dim fStream     As String
Dim fldStream   As String
Dim frmChk      As String
Dim hStream     As String
Dim skip1       As String
Dim flChk       As String
Dim shStream    As String
Dim shareAdd    As Boolean

With EReplicator
    .FLocation.Text = ""
    .FileList.Clear
    .DestFolder.Text = ""
    .List1.Clear
    .ManualAdd.Text = ""
    .ComChar.Text = ""
    .SConn.Clear
    .FConn.Clear
    .HCount.Caption = "0"
End With

Open FileToOpen For Input As #1
Line Input #1, skip1
If Not skip1 = "[Files]" Then
    MsgBox "This is not a valid session file, or file is corrupted!", vbCritical, "Error:"
    Close #1
    EReplicator.Caption = "Replicator 1-2-3 (Untitled.sss)"
    Exit Function
End If

Do While Not EOF(1)
    Line Input #1, fStream
    If fStream = "[Share]" Then Exit Do
    If fStream = "[Folder]" Then MsgBox "File Format is " & _
                 "Invalid.", vbCritical, "Error:": Close #1: Exit Function
    EReplicator.FileList.AddItem Trim(fStream)
Loop

Line Input #1, shStream
EReplicator.RShare.Text = Trim(shStream)
shareAdd = True
s = 0
Do
    If EReplicator.RShare.List(s) = Trim(shStream) Then shareAdd = False
    s = s + 1
Loop Until s > EReplicator.RShare.ListCount - 1
If shareAdd Then EReplicator.RShare.AddItem (shStream)

Line Input #1, flChk
If Not flChk = "[Folder]" Then MsgBox "File Format is " & _
                "Invalid.", vbCritical, "Error:": Close #1: Exit Function
Line Input #1, fldStream
EReplicator.DestFolder.Text = Trim(fldStream)
Line Input #1, frmChk
If Not frmChk = "[Hosts]" Then MsgBox "File Format is " & _
                "Invalid.", vbCritical, "Error:": Close #1: Exit Function

Do While Not EOF(1)
   Line Input #1, hStream
   EReplicator.List1.AddItem Trim(hStream)
Loop
Close #1

I = 0
Do
    If EReplicator.List1.List(I) = "" Then
        EReplicator.List1.RemoveItem (I)
    Else
        I = I + 1
    End If
Loop Until I > EReplicator.List1.ListCount - 1

J = 0
Do
    If EReplicator.FileList.List(J) = "" Then
        EReplicator.FileList.RemoveItem (J)
    Else
        J = J + 1
    End If
Loop Until J > EReplicator.FileList.ListCount - 1

EReplicator.HCount.Caption = EReplicator.List1.ListCount
End Function

Public Function SaveFile(FileToSave As String)
Dim I       As Integer
Dim J       As Integer

Open FileToSave For Output As #1
Print #1, "[Files]"
I = 0
Do
    Print #1, EReplicator.FileList.List(I)
    I = I + 1
Loop Until I > EReplicator.FileList.ListCount - 1
Print #1, "[Share]"
Print #1, EReplicator.RShare.Text
Print #1, "[Folder]"
Print #1, EReplicator.DestFolder.Text
Print #1, "[Hosts]"
J = 0
Do
    Print #1, EReplicator.List1.List(J)
    J = J + 1
Loop Until J > EReplicator.List1.ListCount - 1
Close #1
End Function

Public Function PlaySound(ResourceId As Integer)
Dim FileData() As Byte
FileData = LoadResData(ResourceId, "CUSTOM")
sndPlaySound FileData(LBound(FileData)), _
    SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY
DoEvents
End Function
