Attribute VB_Name = "mdlSelectDirectory"
Option Explicit
Private Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type
Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const BIF_DONTGOBELOWDOMAIN = &H2
Private Const BIF_STATUSTEXT = &H4
Private Const BIF_RETURNFSANCESTORS = &H8
Private Const BIF_BROWSEFORCOMPUTER = &H1000
Private Const BIF_BROWSEFORPRINTER = &H2000
Private Const MAX_PATH = 255
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)

'Public Sub ProcessRuntimeError(lNumber As Long, lDescription As String, lSub As String)
'lErrHandler.ProcessError lNumber, lDescription, lSub
'Err.Clear
'Exit Sub
'ErrHandler:
 '   Err.Clear
'    'ProcessRuntimeError Err.Number, Err.Description, "Public Sub ProcessRuntimeError(lNumber As Long, lDescription As String, lSub As String)"
'End Sub

Public Function SelectDirectory(lForm As Form, Optional lTitle As String = "Select Folder") As String
On Local Error Resume Next
Dim b As BROWSEINFO, l As Long, lPath As String, i As Integer ', MAX_lPath
With b
    .hOwner = lForm.hWnd
    .pidlRoot = 0&
    .lpszTitle = lTitle
    .ulFlags = BIF_RETURNONLYFSDIRS
End With
l = SHBrowseForFolder(b)
lPath = Space$(255)
If SHGetPathFromIDList(ByVal l, ByVal lPath) Then
    i = InStr(lPath, Chr$(0))
    SelectDirectory = Left(lPath, i - 1)
Else
    SelectDirectory = ""
End If
Call CoTaskMemFree(l)
End Function
