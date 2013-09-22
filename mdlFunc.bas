Attribute VB_Name = "mdlFunc"
Option Explicit
Public Enum e_DriveListOpts
    OPT_CDWRITERS
    OPT_DVD
    OPT_ALL
End Enum
Public Type t_AudioTrack
    Album As String
    Artist As String
    Title As String
    no As Integer
    grab As Boolean
    startLBA As Long
    endLBA As Long
    lenLBA As Long
End Type
Public Type t_AudioTracks
    Track(98) As t_AudioTrack
    count As Integer
End Type
Public cManager As New FL_Manager
Public strDrvID As String

Public Function GetFileTitle(lFilename As String) As String
On Local Error Resume Next
Dim msg() As String
If Len(lFilename) <> 0 Then
    msg = Split(lFilename, "\", -1, vbTextCompare)
    GetFileTitle = msg(UBound(msg))
End If
End Function

Public Function DoesFileExist(lFilename As String) As Boolean
On Local Error Resume Next
Dim msg As String
msg = Dir(lFilename)
If msg <> "" Then
    DoesFileExist = True
Else
    DoesFileExist = False
End If
End Function

Public Function PathFromPathFile(ByVal strText As String) As String
On Error Resume Next
PathFromPathFile = Left$(strText, InStrRev(strText, "\"))
End Function

Public Function FileFromPath(ByVal strText As String) As String
On Error Resume Next
FileFromPath = Mid$(strText, InStrRev(strText, "\") + 1)
End Function

Public Function AddSlash(ByVal strText As String) As String
If InStr(strText, "/") > 0 Then
    If Not Right$(strText, 1) = "/" Then strText = strText & "/"
Else
    If Not Right$(strText, 1) = "\" Then strText = strText & "\"
End If
AddSlash = strText
End Function

Public Function GetDriveList(options As e_DriveListOpts) As String()
    Dim cDrvNfo     As New FL_DriveInfo
    Dim strDrvs()   As String
    Dim strRet()    As String
    ReDim strRet(0) As String
    Dim i           As Integer
    strDrvs = cManager.GetCDVDROMs
    For i = LBound(strDrvs) To UBound(strDrvs) - 1
        cDrvNfo.GetInfo cManager.DrvChr2DrvID(strDrvs(i))
        Select Case options
            Case OPT_ALL:
                strRet(UBound(strRet)) = strDrvs(i)
                ReDim Preserve strRet(UBound(strRet) + 1) As String
            Case OPT_DVD:
                If (cDrvNfo.ReadCapabilities And RC_DVDROM) Then
                    strRet(UBound(strRet)) = strDrvs(i)
                    ReDim Preserve strRet(UBound(strRet) + 1) As String
                End If
            Case OPT_CDWRITERS:
                If (cDrvNfo.WriteCapabilities And WC_CDR) Then
                    strRet(UBound(strRet)) = strDrvs(i)
                    ReDim Preserve strRet(UBound(strRet) + 1) As String
                End If
        End Select
    Next i
    GetDriveList = strRet
End Function
