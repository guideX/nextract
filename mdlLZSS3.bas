Attribute VB_Name = "mdlLZSS3"
Option Explicit
Public Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public OriginalArray() As Byte
Public OriginalSize As Long
Public Const DictionarySize = 16
Private Type gStream
    sData() As Byte
    sPosition As Long
    sBitPosition As Byte
    sBuffer As Byte
End Type
Private lStream(3) As gStream
Private MaxHistory As Long

Public Sub CompressLZSS3(ByteArray() As Byte)
On Local Error Resume Next
Dim l As Long, o As Long, n As Long, i As Integer, g As Long, l2 As Long, o2 As Long, n2 As Long, g2 As Long, msg As String, msg2 As String, msg3 As String, cur As Currency, t As Integer
cur = UBound(ByteArray)
t = 0
Call BeginLZSS
MaxHistory = CLng(1024) * DictionarySize
msg = StrConv(ByteArray(), vbUnicode)
Call AddBitsToStream(lStream(3), CByte(DictionarySize), 8)
Call AddBitsToStream(lStream(3), ByteArray(0), 8)
l = 2
Do While l + 2 <= UBound(ByteArray)
    t = Int((l / cur) * 100)
    frmMain.FileProgressLbl.Caption = t & "%"
    frmMain.prgMakeProgress.Value = t
    DoEvents
    i = 2
    o = l - 1 - MaxHistory
    If o < 1 Then o = 1
    Do
        n = o
        i = i + 1
        If i = 259 Then Exit Do
        msg2 = Mid(msg, l, i)
        If Len(msg2) < i Then Exit Do
        o = InStr(n, msg, msg2)
    Loop While o <> l
    i = i - 1
    If i < 3 Then
        Call AddBitsToStream(lStream(0), 0, 1)
        Call AddBitsToStream(lStream(3), ByteArray(l - 1), 8)
        l = l + 1
    Else
        o = l - 1 - MaxHistory
        If o < 1 Then o = 1
        n = l - (InStrRev(Mid(msg, o, l - o + i - 1), Mid(msg, l, i)) + o - 1)
        Call AddBitsToStream(lStream(0), 1, 1)
        Call AddBitsToStream(lStream(2), i - 3, 8)
        Call AddBitsToStream(lStream(1), (n And &HFF00) / &H100, 8)
        Call AddBitsToStream(lStream(1), n And &HFF, 8)
        l = l + i
    End If
Loop
If l - 1 <= UBound(ByteArray) Then
    For n2 = l - 1 To UBound(ByteArray)
        Call AddBitsToStream(lStream(0), 0, 1)
        Call AddBitsToStream(lStream(3), ByteArray(n2), 8)
    Next
End If
Call AddBitsToStream(lStream(0), 1, 1)
Call AddBitsToStream(lStream(1), 0, 8)
Call AddBitsToStream(lStream(1), 0, 8)
For n2 = 0 To 3
    Do While lStream(n2).sBitPosition > 0
        Call AddBitsToStream(lStream(n2), 0, 1)
    Loop
Next n2
o2 = 0
For n2 = 0 To 3
    If lStream(n2).sPosition > 0 Then
        ReDim Preserve lStream(n2).sData(lStream(n2).sPosition - 1)
        o2 = o2 + lStream(n2).sPosition
    Else
        ReDim lStream(n2).sData(0)
        o2 = o2 + 1
    End If
Next n2
ReDim ByteArray(o2 + 5)
ByteArray(0) = Int(UBound(lStream(0).sData) / &H10000) And &HFF
ByteArray(1) = Int(UBound(lStream(0).sData) / &H100) And &HFF
ByteArray(2) = UBound(lStream(0).sData) And &HFF
ByteArray(3) = Int(UBound(lStream(2).sData) / &H10000) And &HFF
ByteArray(4) = Int(UBound(lStream(2).sData) / &H100) And &HFF
ByteArray(5) = UBound(lStream(2).sData) And &HFF
l = 6
For n2 = 0 To 3
    For g2 = 0 To UBound(lStream(n2).sData)
        ByteArray(l) = lStream(n2).sData(g2)
        l = l + 1
    Next
Next n2
frmMain.FileProgressLbl.Caption = "100%"
frmMain.prgMakeProgress.Value = 100
DoEvents
End Sub

Private Sub BeginLZSS()
On Local Error Resume Next
Dim i As Integer
For i = 0 To 3
    ReDim lStream(i).sData(10)
    lStream(i).sBitPosition = 0
    lStream(i).sBuffer = 0
    lStream(i).sPosition = 0
Next
End Sub

Private Sub AddBitsToStream(Toarray As gStream, Number As Byte, Numbits As Byte)
On Local Error Resume Next
Dim l As Long
If Numbits = 8 And Toarray.sBitPosition = 0 Then
    If Toarray.sPosition > UBound(Toarray.sData) Then ReDim Preserve Toarray.sData(Toarray.sPosition + 500)
    Toarray.sData(Toarray.sPosition) = Number And &HFF
    Toarray.sPosition = Toarray.sPosition + 1
    Exit Sub
End If
For l = Numbits - 1 To 0 Step -1
    Toarray.sBuffer = Toarray.sBuffer * 2 + (-1 * ((Number And 2 ^ l) > 0))
    Toarray.sBitPosition = Toarray.sBitPosition + 1
    If Toarray.sBitPosition = 8 Then
        If Toarray.sPosition > UBound(Toarray.sData) Then ReDim Preserve Toarray.sData(Toarray.sPosition + 500)
        Toarray.sData(Toarray.sPosition) = Toarray.sBuffer
        Toarray.sBitPosition = 0
        Toarray.sBuffer = 0
        Toarray.sPosition = Toarray.sPosition + 1
    End If
Next l
End Sub

Private Function ReadBitsFromArray(FromArray() As Byte, FromPos As Long, FromBit As Byte, Numbits As Integer) As Long
On Local Error Resume Next
Dim i As Integer, l As Long
If FromBit = 0 And Numbits = 8 Then
    ReadBitsFromArray = FromArray(FromPos)
    FromPos = FromPos + 1
    Exit Function
End If
For i = 1 To Numbits
    l = l * 2 + (-1 * ((FromArray(FromPos) And 2 ^ (7 - FromBit)) > 0))
    FromBit = FromBit + 1
    If FromBit = 8 Then
        If FromPos + 1 > UBound(FromArray) Then
            Do While i < Numbits
                l = l * 2
                i = i + 1
            Loop
            FromPos = FromPos + 1
            Exit For
        End If
        FromPos = FromPos + 1
        FromBit = 0
    End If
Next
ReadBitsFromArray = l
End Function
