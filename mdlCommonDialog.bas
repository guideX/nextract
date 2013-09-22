Attribute VB_Name = "mdlCommonDialog"
Option Explicit
Private Type gPlayerFileTypes
    pWindowsMediaFormats As String
    pVFMp3PlayerFormats As String
End Type
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
    End Type
    Public Const OFN_READONLY = &H1
    Public Const OFN_OVERWRITEPROMPT = &H2
    Public Const OFN_HIDEREADONLY = &H4
    Public Const OFN_NOCHANGEDIR = &H8
    Public Const OFN_SHOWHELP = &H10
    Public Const OFN_ENABLEHOOK = &H20
    Public Const OFN_ENABLETEMPLATE = &H40
    Public Const OFN_ENABLETEMPLATEHANDLE = &H80
    Public Const OFN_NOVALIDATE = &H100
    Public Const OFN_ALLOWMULTISELECT = &H200
    Public Const OFN_EXTENSIONDIFFERENT = &H400
    Public Const OFN_PATHMUSTEXIST = &H800
    Public Const OFN_FILEMUSTEXIST = &H1000
    Public Const OFN_CREATEPROMPT = &H2000
    Public Const OFN_SHAREAWARE = &H4000
    Public Const OFN_NOREADONLYRETURN = &H8000
    Public Const OFN_NOTESTFILECREATE = &H10000
    Public Const OFN_NONETWORKBUTTON = &H20000
    Public Const OFN_NOLONGNAMES = &H40000
    Public Const OFN_EXPLORER = &H80000
    Public Const OFN_NODEREFERENCELINKS = &H100000
    Public Const OFN_LONGNAMES = &H200000
    Public Const OFN_SHAREFALLTHROUGH = 2
    Public Const OFN_SHARENOWARN = 1
    Public Const OFN_SHAREWARN = 0
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Function SaveDialog(Form1 As Form, Filter As String, Title As String, InitDir As String) As String
On Local Error Resume Next
Dim ofn As OPENFILENAME, A As Long
ofn.lStructSize = Len(ofn)
ofn.hwndOwner = Form1.hWnd
ofn.hInstance = App.hInstance
If Right$(Filter, 1) <> "|" Then Filter = Filter + "|"
For A = 1 To Len(Filter)
    If Mid$(Filter, A, 1) = "|" Then Mid$(Filter, A, 1) = Chr$(0)
Next A
ofn.lpstrFilter = Filter
ofn.lpstrFile = Space$(254)
ofn.nMaxFile = 255
ofn.lpstrFileTitle = Space$(254)
ofn.nMaxFileTitle = 255
ofn.lpstrInitialDir = InitDir
ofn.lpstrTitle = Title
ofn.Flags = OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT Or OFN_CREATEPROMPT
A = GetSaveFileName(ofn)
If (A) Then
    ofn.lpstrFile = Trim$(ofn.lpstrFile)
    ofn.lpstrFile = Left(ofn.lpstrFile, Len(ofn.lpstrFile) - 1)
    'MsgBox "!" & ofn.lpstrFile & "!"
    SaveDialog = Trim$(ofn.lpstrFile)
Else
    SaveDialog = ""
End If
End Function

Public Function OpenDialog(lForm As Form, lFilter As String, lTitle As String, lInitDir As String) As String
On Local Error Resume Next
Dim ofn As OPENFILENAME, A As Long
ofn.lStructSize = Len(ofn)
ofn.hwndOwner = lForm.hWnd
ofn.hInstance = App.hInstance
If Right$(lFilter, 1) <> "|" Then lFilter = lFilter + "|"
For A = 1 To Len(lFilter)
    If Mid$(lFilter, A, 1) = "|" Then Mid$(lFilter, A, 1) = Chr$(0)
Next A
ofn.lpstrFilter = lFilter
ofn.lpstrFile = Space$(254)
ofn.nMaxFile = 255
ofn.lpstrFileTitle = Space$(254)
ofn.nMaxFileTitle = 255
ofn.lpstrInitialDir = lInitDir
ofn.lpstrTitle = lTitle
ofn.Flags = OFN_HIDEREADONLY Or OFN_FILEMUSTEXIST
A = GetOpenFileName(ofn)
If (A) Then
    Dim msg As String
    msg = Trim$(ofn.lpstrFile)
    OpenDialog = Left(msg, Len(msg) - 1)
Else
    OpenDialog = ""
End If
End Function

