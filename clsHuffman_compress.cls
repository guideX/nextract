VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHuffman"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const PROGRESS_CALCFREQUENCY = 7
Private Const PROGRESS_CALCCRC = 5
Private Const PROGRESS_ENCODING = 88
Event Progress(Procent As Integer)
Private Type HUFFMANTREE
    ParentNode As Integer
    RightNode As Integer
    LeftNode As Integer
    Value As Integer
    Weight As Long
End Type
Private Type ByteArray
    count As Byte
    Data() As Byte
End Type
Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Sub EncodeFile(SourceFile As String, DestFile As String)
Dim ByteArray() As Byte, Filenr As Integer
If (Not FileExist(SourceFile)) Then Err.Raise vbObjectError, "clsHuffman.EncodeFile()", "Source file does not exist"
Filenr = FreeFile
Open SourceFile For Binary As #Filenr
    ReDim ByteArray(0 To LOF(Filenr) - 1)
    Get #Filenr, , ByteArray()
Close #Filenr
Call EncodeByte(ByteArray(), UBound(ByteArray) + 1)
If (FileExist(DestFile)) Then Kill DestFile
Open DestFile For Binary As #Filenr
    Put #Filenr, , ByteArray()
Close #Filenr
End Sub

Private Sub CreateTree(Nodes() As HUFFMANTREE, NodesCount As Long, Char As Long, Bytes As ByteArray)
On Local Error Resume Next
Dim A As Integer, NodeIndex As Long, msg As String, Title As String
NodeIndex = 0
For A = 0 To (Bytes.count - 1)
    If (Bytes.Data(A) = 0) Then
        If (Nodes(NodeIndex).LeftNode = -1) Then
            Nodes(NodeIndex).LeftNode = NodesCount
            Nodes(NodesCount).ParentNode = NodeIndex
            Nodes(NodesCount).LeftNode = -1
            Nodes(NodesCount).RightNode = -1
            Nodes(NodesCount).Value = -1
            NodesCount = NodesCount + 1
        End If
        NodeIndex = Nodes(NodeIndex).LeftNode
    ElseIf (Bytes.Data(A) = 1) Then
        If (Nodes(NodeIndex).RightNode = -1) Then
            Nodes(NodeIndex).RightNode = NodesCount
            Nodes(NodesCount).ParentNode = NodeIndex
            Nodes(NodesCount).LeftNode = -1
            Nodes(NodesCount).RightNode = -1
            Nodes(NodesCount).Value = -1
            NodesCount = NodesCount + 1
        End If
        NodeIndex = Nodes(NodeIndex).RightNode
    Else
        frmMain.ShowMessage "Packing could not be completed", Err.Number
        Close
        End
    End If
Next A
Nodes(NodeIndex).Value = Char
End Sub

Public Sub EncodeByte(ByteArray() As Byte, ByteLen As Long)
On Local Error Resume Next
Dim i As Long, j As Long, Char As Byte, BitPos As Byte, lNode1 As Long, lNode2 As Long, lNodes As Long, lLength As Long, count As Integer, lWeight1 As Long, lWeight2 As Long, Result() As Byte, ByteValue As Byte, ResultLen As Long, Bytes As ByteArray, NodesCount As Integer, NewProgress As Integer, CurrProgress As Integer, BitValue(0 To 7) As Byte, CharCount(0 To 255) As Long, Nodes(0 To 511) As HUFFMANTREE, CharValue(0 To 255) As ByteArray
If (ByteLen = 0) Then
    ReDim Preserve ByteArray(0 To ByteLen + 3)
    If (ByteLen > 0) Then
        Call CopyMem(ByteArray(4), ByteArray(0), ByteLen)
    End If
    ByteArray(0) = 72
    ByteArray(1) = 69
    ByteArray(2) = 48
    ByteArray(3) = 13
    Exit Sub
End If
ReDim Result(0 To 522)
Result(0) = 72
Result(1) = 69
Result(2) = 51
Result(3) = 13
ResultLen = 4
For i = 0 To (ByteLen - 1)
    CharCount(ByteArray(i)) = CharCount(ByteArray(i)) + 1
    If (i Mod 1000 = 0) Then
        NewProgress = i / ByteLen * PROGRESS_CALCFREQUENCY
        If (NewProgress <> CurrProgress) Then
            CurrProgress = NewProgress
            RaiseEvent Progress(CurrProgress)
        End If
    End If
Next i
For i = 0 To 255
    If (CharCount(i) > 0) Then
        With Nodes(NodesCount)
            .Weight = CharCount(i)
            .Value = i
            .LeftNode = -1
            .RightNode = -1
            .ParentNode = -1
        End With
        NodesCount = NodesCount + 1
    End If
Next i
For lNodes = NodesCount To 2 Step -1
    lNode1 = -1: lNode2 = -1
    For i = 0 To (NodesCount - 1)
        If (Nodes(i).ParentNode = -1) Then
            If (lNode1 = -1) Then
                lWeight1 = Nodes(i).Weight
                lNode1 = i
            ElseIf (lNode2 = -1) Then
                lWeight2 = Nodes(i).Weight
                lNode2 = i
            ElseIf (Nodes(i).Weight < lWeight1) Then
                If (Nodes(i).Weight < lWeight2) Then
                    If (lWeight1 < lWeight2) Then
                        lWeight2 = Nodes(i).Weight
                        lNode2 = i
                    Else
                        lWeight1 = Nodes(i).Weight
                        lNode1 = i
                    End If
                Else
                    lWeight1 = Nodes(i).Weight
                    lNode1 = i
                End If
            ElseIf (Nodes(i).Weight < lWeight2) Then
                lWeight2 = Nodes(i).Weight
                lNode2 = i
            End If
        End If
    Next
    With Nodes(NodesCount)
        .Weight = lWeight1 + lWeight2
        .LeftNode = lNode1
        .RightNode = lNode2
        .ParentNode = -1
        .Value = -1
    End With
    Nodes(lNode1).ParentNode = NodesCount
    Nodes(lNode2).ParentNode = NodesCount
    NodesCount = NodesCount + 1
Next lNodes
ReDim Bytes.Data(0 To 255)
Call CreateBitSequences(Nodes(), NodesCount - 1, Bytes, CharValue)
For i = 0 To 255
    If (CharCount(i) > 0) Then
        lLength = lLength + CharValue(i).count * CharCount(i)
    End If
Next i
lLength = IIf(lLength Mod 8 = 0, lLength \ 8, lLength \ 8 + 1)
If ((lLength = 0) Or (lLength > ByteLen)) Then
    ReDim Preserve ByteArray(0 To ByteLen + 3)
    Call CopyMem(ByteArray(4), ByteArray(0), ByteLen)
    ByteArray(0) = 72
    ByteArray(1) = 69
    ByteArray(2) = 48
    ByteArray(3) = 13
    Exit Sub
End If
Char = 0
For i = 0 To (ByteLen - 1)
    Char = Char Xor ByteArray(i)
    If (i Mod 10000 = 0) Then
        NewProgress = i / ByteLen * PROGRESS_CALCCRC + PROGRESS_CALCFREQUENCY
        If (NewProgress <> CurrProgress) Then
            CurrProgress = NewProgress
            RaiseEvent Progress(CurrProgress)
        End If
    End If
Next i
Result(ResultLen) = Char
ResultLen = ResultLen + 1
Call CopyMem(Result(ResultLen), ByteLen, 4)
ResultLen = ResultLen + 4
For i = 0 To 7
    BitValue(i) = 2 ^ i
Next i
count = 0
For i = 0 To 255
  If (CharValue(i).count > 0) Then
    count = count + 1
  End If
Next
Call CopyMem(Result(ResultLen), count, 2)
ResultLen = ResultLen + 2
count = 0
For i = 0 To 255
  If (CharValue(i).count > 0) Then
    Result(ResultLen) = i
    ResultLen = ResultLen + 1
    Result(ResultLen) = CharValue(i).count
    ResultLen = ResultLen + 1
    count = count + 16 + CharValue(i).count
  End If
Next
ReDim Preserve Result(0 To ResultLen + count \ 8)
BitPos = 0
ByteValue = 0
For i = 0 To 255
  With CharValue(i)
    If (.count > 0) Then
      For j = 0 To (.count - 1)
        If (.Data(j)) Then ByteValue = ByteValue + BitValue(BitPos)
        BitPos = BitPos + 1
        If (BitPos = 8) Then
          Result(ResultLen) = ByteValue
          ResultLen = ResultLen + 1
          ByteValue = 0
          BitPos = 0
        End If
      Next
    End If
  End With
Next
If (BitPos > 0) Then
  Result(ResultLen) = ByteValue
  ResultLen = ResultLen + 1
End If
ReDim Preserve Result(0 To ResultLen - 1 + lLength)
Char = 0
BitPos = 0
For i = 0 To (ByteLen - 1)
    With CharValue(ByteArray(i))
        For j = 0 To (.count - 1)
            If (.Data(j) = 1) Then Char = Char + BitValue(BitPos)
            BitPos = BitPos + 1
            If (BitPos = 8) Then
                Result(ResultLen) = Char
                ResultLen = ResultLen + 1
                BitPos = 0
                Char = 0
            End If
        Next j
    End With
    If (i Mod 10000 = 0) Then
        NewProgress = i / ByteLen * PROGRESS_ENCODING + PROGRESS_CALCCRC + PROGRESS_CALCFREQUENCY
        If (NewProgress <> CurrProgress) Then
            CurrProgress = NewProgress
            RaiseEvent Progress(CurrProgress)
        End If
    End If
Next i
If (BitPos > 0) Then
    Result(ResultLen) = Char
    ResultLen = ResultLen + 1
End If
ReDim ByteArray(0 To ResultLen - 1)
Call CopyMem(ByteArray(0), Result(0), ResultLen)
If (CurrProgress <> 100) Then RaiseEvent Progress(100)
End Sub

Public Function EncodeString(Text As String) As String
Dim ByteArray() As Byte
ByteArray() = StrConv(Text, vbFromUnicode)
Call EncodeByte(ByteArray, Len(Text))
EncodeString = StrConv(ByteArray(), vbUnicode)
End Function

Private Sub CreateBitSequences(Nodes() As HUFFMANTREE, ByVal NodeIndex As Integer, Bytes As ByteArray, CharValue() As ByteArray)
Dim NewBytes As ByteArray
If (Nodes(NodeIndex).Value > -1) Then
    CharValue(Nodes(NodeIndex).Value) = Bytes
    Exit Sub
End If
If (Nodes(NodeIndex).LeftNode > -1) Then
    NewBytes = Bytes
    NewBytes.Data(NewBytes.count) = 0
    NewBytes.count = NewBytes.count + 1
    Call CreateBitSequences(Nodes(), Nodes(NodeIndex).LeftNode, NewBytes, CharValue)
End If
If (Nodes(NodeIndex).RightNode > -1) Then
    NewBytes = Bytes
    NewBytes.Data(NewBytes.count) = 1
    NewBytes.count = NewBytes.count + 1
    Call CreateBitSequences(Nodes(), Nodes(NodeIndex).RightNode, NewBytes, CharValue)
End If
End Sub

Private Function FileExist(FileName As String) As Boolean
Call FileLen(FileName)
FileExist = True
Exit Function
FileDoesNotExist:
FileExist = False
End Function
