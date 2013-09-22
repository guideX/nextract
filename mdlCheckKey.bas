Attribute VB_Name = "mdlCheckKey"
Option Explicit
Private aDecTab(255) As Integer
Private aEncTab(63) As Byte

Private Function PadString(strData As String) As String
On Local Error Resume Next
Dim nLen As Long, sPad As String, nPad As Integer
nLen = Len(strData)
nPad = ((nLen \ 8) + 1) * 8 - nLen
sPad = String(nPad, Chr(nPad))
PadString = strData & sPad
'If Err.Number <> 0 Then 'ProcessRuntimeError Err.Number, Err.Description, "Private Function PadString(strData As String) As String"
End Function

Private Function UnpadString(strData As String) As String
On Local Error Resume Next
Dim nLen As Long, nPad As Long
nLen = Len(strData)
If nLen = 0 Then Exit Function
nPad = Asc(Right(strData, 1))
If nPad > 8 Then nPad = 0
UnpadString = Left(strData, nLen - nPad)
'If Err.Number <> 0 Then 'ProcessRuntimeError Err.Number, Err.Description, "Private Function UnpadString(strData As String) As String"
End Function

Public Function EncodeStr64(encString As String, ByVal MaxPerLine As Integer) As String
On Local Error Resume Next
Dim abOutput() As Byte, sLast As String, b(3) As Byte, j As Integer, CharCount As Integer, iIndex As Long, Umax As Long, i As Long, nLen As Long, nQuants As Long
EncodeStr64 = ""
nLen = Len(encString)
nQuants = nLen \ 3
iIndex = 0
If MaxPerLine < 10 Then MaxPerLine = 10
Umax = nQuants + 1
Call MakeEncTab
If (nQuants > 0) Then
    ReDim abOutput(nQuants * 4 - 1)
    For i = 0 To nQuants - 1
        For j = 0 To 2
            b(j) = Asc(Mid(encString, (i * 3) + j + 1, 1))
        Next
        Call EncodeQuantumB(b)
        abOutput(iIndex) = b(0)
        abOutput(iIndex + 1) = b(1)
        abOutput(iIndex + 2) = b(2)
        abOutput(iIndex + 3) = b(3)
        CharCount = CharCount + 4
        If CharCount >= MaxPerLine Then
            ReDim Preserve abOutput(UBound(abOutput) + 2)
            CharCount = 0
            abOutput(iIndex + 4) = 13
            abOutput(iIndex + 5) = 10
            iIndex = iIndex + 6
            Else
            iIndex = iIndex + 4
        End If
    Next i
    EncodeStr64 = StrConv(abOutput, vbUnicode)
End If
Select Case nLen Mod 3
Case 0
    sLast = ""
Case 1
    b(0) = Asc(Mid(encString, nLen, 1))
    b(1) = 0
    b(2) = 0
    Call EncodeQuantumB(b)
    sLast = StrConv(b(), vbUnicode)
    sLast = Left(sLast, 2) & "=="
Case 2
    b(0) = Asc(Mid(encString, nLen - 1, 1))
    b(1) = Asc(Mid(encString, nLen, 1))
    b(2) = 0
    Call EncodeQuantumB(b)
    sLast = StrConv(b(), vbUnicode)
    sLast = Left(sLast, 3) & "="
End Select
EncodeStr64 = EncodeStr64 & sLast
'If Err.Number <> 0 Then 'ProcessRuntimeError Err.Number, Err.Description, "Public Function EncodeStr64(encString As String, ByVal MaxPerLine As Integer) As String"
End Function

Public Function DecodeStr64(decString As String) As String
On Local Error Resume Next
Dim abDecoded() As Byte, d(3) As Byte, c As Integer, di As Integer, i As Long, nLen As Long, iIndex As Long, Umax As Long
nLen = Len(decString)
If nLen < 4 Then Exit Function
ReDim abDecoded(((nLen \ 4) * 3) - 1)
Umax = nLen
iIndex = 0
di = 0
Call MakeDecTab
For i = 1 To Len(decString)
    c = CByte(Asc(Mid(decString, i, 1)))
    c = aDecTab(c)
    If c >= 0 Then
        d(di) = CByte(c)
        di = di + 1
        If di = 4 Then
            abDecoded(iIndex) = SHL2(d(0)) Or (SHR4(d(1)) And &H3)
            iIndex = iIndex + 1
            abDecoded(iIndex) = SHL4(d(1) And &HF) Or (SHR2(d(2)) And &HF)
            iIndex = iIndex + 1
            abDecoded(iIndex) = SHL6(d(2) And &H3) Or d(3)
            iIndex = iIndex + 1
            If d(3) = 64 Then
                iIndex = iIndex - 1
                abDecoded(iIndex) = 0
            End If
            If d(2) = 64 Then
                iIndex = iIndex - 1
                abDecoded(iIndex) = 0
            End If
            di = 0
        End If
    End If
Next i
DecodeStr64 = StrConv(abDecoded(), vbUnicode)
DecodeStr64 = Left(DecodeStr64, iIndex)
'If Err.Number <> 0 Then 'ProcessRuntimeError Err.Number, Err.Description, "Public Function DecodeStr64(decString As String) As String"
End Function

Private Sub EncodeQuantumB(b() As Byte)
On Local Error Resume Next
Dim b0 As Byte, b1 As Byte, b2 As Byte, b3 As Byte
b0 = SHR2(b(0)) And &H3F
b1 = SHL4(b(0) And &H3) Or (SHR4(b(1)) And &HF)
b2 = SHL2(b(1) And &HF) Or (SHR6(b(2)) And &H3)
b3 = b(2) And &H3F
b(0) = aEncTab(b0)
b(1) = aEncTab(b1)
b(2) = aEncTab(b2)
b(3) = aEncTab(b3)
'If Err.Number <> 0 Then 'ProcessRuntimeError Err.Number, Err.Description, "Private Sub EncodeQuantumB(b() As Byte)"
End Sub

Private Function MakeDecTab()
On Local Error Resume Next
Dim t As Integer
Dim c As Integer
For c = 0 To 255
    aDecTab(c) = -1
Next
t = 0
For c = Asc("A") To Asc("Z")
    aDecTab(c) = t
    t = t + 1
Next
For c = Asc("a") To Asc("z")
    aDecTab(c) = t
    t = t + 1
Next
For c = Asc("0") To Asc("9")
    aDecTab(c) = t
    t = t + 1
Next
c = Asc("+")
aDecTab(c) = t
t = t + 1
c = Asc("/")
aDecTab(c) = t
t = t + 1
c = Asc("=")
aDecTab(c) = t
'If Err.Number <> 0 Then 'ProcessRuntimeError Err.Number, Err.Description, "Private Function MakeDecTab()"
End Function

Private Function MakeEncTab()
On Local Error Resume Next
Dim i As Integer
Dim c As Integer
i = 0
For c = Asc("A") To Asc("Z")
    aEncTab(i) = c
    i = i + 1
Next
For c = Asc("a") To Asc("z")
    aEncTab(i) = c
    i = i + 1
Next
For c = Asc("0") To Asc("9")
    aEncTab(i) = c
    i = i + 1
Next
c = Asc("+")
aEncTab(i) = c
i = i + 1
c = Asc("/")
aEncTab(i) = c
i = i + 1
'If Err.Number <> 0 Then 'ProcessRuntimeError Err.Number, Err.Description, "Private Function MakeEncTab()"
End Function

Private Function SHL2(ByVal bytValue As Byte) As Byte
On Local Error Resume Next
SHL2 = (bytValue * &H4) And &HFF
'If Err.Number <> 0 Then 'ProcessRuntimeError Err.Number, Err.Description, "Private Function SHL2(ByVal bytValue As Byte) As Byte"
End Function

Private Function SHL4(ByVal bytValue As Byte) As Byte
On Local Error Resume Next
SHL4 = (bytValue * &H10) And &HFF
'If Err.Number <> 0 Then 'ProcessRuntimeError Err.Number, Err.Description, "Private Function SHL4(ByVal bytValue As Byte) As Byte"
End Function

Private Function SHL6(ByVal bytValue As Byte) As Byte
On Local Error Resume Next
SHL6 = (bytValue * &H40) And &HFF
'If Err.Number <> 0 Then 'ProcessRuntimeError Err.Number, Err.Description, "Private Function SHL6(ByVal bytValue As Byte) As Byte"
End Function

Private Function SHR2(ByVal bytValue As Byte) As Byte
On Local Error Resume Next
SHR2 = bytValue \ &H4
'If Err.Number <> 0 Then 'ProcessRuntimeError Err.Number, Err.Description, "Private Function SHR2(ByVal bytValue As Byte) As Byte"
End Function

Private Function SHR4(ByVal bytValue As Byte) As Byte
On Local Error Resume Next
SHR4 = bytValue \ &H10
'If Err.Number <> 0 Then 'ProcessRuntimeError Err.Number, Err.Description, "Private Function SHR4(ByVal bytValue As Byte) As Byte"
End Function

Private Function SHR6(ByVal bytValue As Byte) As Byte
On Local Error Resume Next
SHR6 = bytValue \ &H40
'If Err.Number <> 0 Then 'ProcessRuntimeError Err.Number, Err.Description, "Private Function SHR6(ByVal bytValue As Byte) As Byte"
End Function

Public Function KeyQuality(ByVal aKey As String) As Integer
On Local Error Resume Next
Dim QC As Integer, LN As Integer, k As Integer, Uc As Boolean, Lc As Boolean, Wid As Integer, ValidKey As Boolean
LN = Len(aKey)
QC = LN * 4
If Len(aKey) < 5 Then KeyQuality = 0: Exit Function
For Wid = 1 To Int(Len(aKey) / 2)
    ValidKey = False
    For k = Wid + 1 To Len(aKey) Step Wid
        If Mid(aKey, 1, Wid) <> Mid(aKey, k, Wid) Then ValidKey = True: Exit For
    Next k
    If ValidKey = False Then Exit For
Next Wid
If ValidKey = False Then KeyQuality = 0: Exit Function
For k = 1 To Len(aKey)
    If Asc(Mid(aKey, k, 1)) > 64 And Asc(Mid(aKey, k, 1)) < 91 Then Uc = True
    If Asc(Mid(aKey, k, 1)) > 96 And Asc(Mid(aKey, k, 1)) < 123 Then Lc = True
Next
If Uc = True And Lc = True Then QC = QC * 1.5
For k = 1 To Len(aKey)
    If Asc(Mid(aKey, k, 1)) > 47 And Asc(Mid(aKey, k, 1)) < 58 Then
        If Uc = True Or Lc = True Then QC = QC * 1.5
        Exit For
    End If
Next k
For k = 1 To Len(aKey)
    If Asc(Mid(aKey, k, 1)) < 48 Or Asc(Mid(aKey, k, 1)) > 122 Or (Asc(Mid(aKey, k, 1)) > 57 And Asc(Mid(aKey, k, 1)) < 65) Then QC = QC * 1.5: Exit For
Next k
If QC > 100 Then QC = 100
KeyQuality = Int(QC)
'If Err.Number <> 0 Then 'ProcessRuntimeError Err.Number, Err.Description, "Public Function KeyQuality(ByVal aKey As String) As Integer"
End Function

