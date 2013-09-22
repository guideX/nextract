Option Strict Off
Option Explicit On
Friend Class clsHuffman
	'Huffman Encoding/Decoding Class
	'-------------------------------
	'
	'(c) 2000, Fredrik Qvarfort
	'
	
	
	'Progress Values for the decoding routine
	Private Const PROGRESS_DECODING As Short = 89
	Private Const PROGRESS_CHECKCRC As Short = 11
	
	'Events
	Event Progress(ByRef Procent As Short)
	
	Private Structure HUFFMANTREE
		Dim ParentNode As Short
		Dim RightNode As Short
		Dim LeftNode As Short
		Dim Value As Short
		Dim Weight As Integer
	End Structure
	
	Private Structure ByteArray
		Dim Count As Byte
		Dim Data() As Byte
	End Structure
	
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Private Declare Sub CopyMem Lib "kernel32"  Alias "RtlMoveMemory"(ByRef Destination As Any, ByRef Source As Any, ByVal Length As Integer)
	
	
	Public Sub DecodeFile(ByRef SourceFile As String, ByRef DestFile As String)
		
		Dim ByteArray() As Byte
		Dim Filenr As Short
		
		'Make sure the source file exists
		If (Not FileExist(SourceFile)) Then
			Err.Raise(vbObjectError, "clsHuffman.DecodeFile()", "Source file does not exist")
		End If
		
		'Read the data from the sourcefile
		Filenr = FreeFile
		FileOpen(Filenr, SourceFile, OpenMode.Binary)
		ReDim ByteArray(LOF(Filenr) - 1)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(Filenr, ByteArray)
		FileClose(Filenr)
		
		'Uncompress the data
		Call DecodeByte(ByteArray, UBound(ByteArray) + 1)
		
		'If the destination file exist we need to
		'destroy it because opening it as binary
		'will not clear the old data
		If (FileExist(DestFile)) Then Kill(DestFile)
		
		'Save the destination string
		FileOpen(Filenr, DestFile, OpenMode.Binary)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(Filenr, ByteArray)
		FileClose(Filenr)
		
	End Sub
	'UPGRADE_NOTE: Char was upgraded to Char_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub CreateTree(ByRef Nodes() As HUFFMANTREE, ByRef NodesCount As Integer, ByRef Char_Renamed As Integer, ByRef Bytes As ByteArray)
		
		Dim a As Short
		Dim NodeIndex As Integer
		
		NodeIndex = 0
		Dim Msg, Title As String
		For a = 0 To (Bytes.Count - 1)
			If (Bytes.Data(a) = 0) Then
				'Left node
				If (Nodes(NodeIndex).LeftNode = -1) Then
					Nodes(NodeIndex).LeftNode = NodesCount
					Nodes(NodesCount).ParentNode = NodeIndex
					Nodes(NodesCount).LeftNode = -1
					Nodes(NodesCount).RightNode = -1
					Nodes(NodesCount).Value = -1
					NodesCount = NodesCount + 1
				End If
				NodeIndex = Nodes(NodeIndex).LeftNode
			ElseIf (Bytes.Data(a) = 1) Then 
				'Right node
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
				'Stop ' removed on suggestion of Roger Gilchrist... was in original code
				Msg = "Un-packing could not be completed." & vbCrLf & "Error In CreateTree sub."
				MsgBox(Msg, MsgBoxStyle.OKOnly + MsgBoxStyle.Exclamation, Title)
				FileClose()
				End
			End If
		Next 
		
		Nodes(NodeIndex).Value = Char_Renamed
		
	End Sub
	
	Public Function DecodeString(ByRef Text As String) As String
		
		Dim ByteArray() As Byte
		
		'Convert the string to a byte array
		'UPGRADE_ISSUE: Constant vbFromUnicode was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		'UPGRADE_TODO: Code was upgraded to use System.Text.UnicodeEncoding.Unicode.GetBytes() which may not have the same behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="93DD716C-10E3-41BE-A4A8-3BA40157905B"'
		ByteArray = System.Text.UnicodeEncoding.Unicode.GetBytes(StrConv(Text, vbFromUnicode))
		
		'Compress the byte array
		Call DecodeByte(ByteArray, Len(Text))
		
		'Convert the compressed byte array to a string
		'UPGRADE_ISSUE: Constant vbUnicode was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		DecodeString = StrConv(System.Text.UnicodeEncoding.Unicode.GetString(ByteArray), vbUnicode)
		
	End Function
	
	
	Public Sub DecodeByte(ByRef ByteArray() As Byte, ByRef ByteLen As Integer)
		
		Dim i As Integer
		Dim j As Integer
		Dim POS As Integer
		'UPGRADE_NOTE: Char was upgraded to Char_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Char_Renamed As Byte
		Dim CurrPos As Integer
		Dim Count As Short
		Dim CheckSum As Byte
		Dim Result() As Byte
		Dim BitPos As Short
		Dim NodeIndex As Integer
		Dim ByteValue As Byte
		Dim ResultLen As Integer
		Dim NodesCount As Integer
		Dim lResultLen As Integer
		Dim NewProgress As Short
		Dim CurrProgress As Short
		Dim BitValue(7) As Byte
		Dim Nodes(511) As HUFFMANTREE
		Dim CharValue(255) As ByteArray
		
		If (ByteArray(0) <> 72) Or (ByteArray(1) <> 69) Or (ByteArray(3) <> 13) Then
			'The source did not contain the identification
			'string "HE?" & vbCr where ? is undefined at
			'the moment (does not matter)
		ElseIf (ByteArray(2) = 48) Then 
			'The text is uncompressed, return the substring
			'Decode = Mid$(Text, 5)
			Call CopyMem(ByteArray(0), ByteArray(4), ByteLen - 4)
			ReDim Preserve ByteArray(ByteLen - 5)
			Exit Sub
		ElseIf (ByteArray(2) <> 51) Then 
			'This is not a Huffman encoded string
			Err.Raise(vbObjectError, "HuffmanDecode()", "The data either was not compressed with HE3 or is corrupt (identification string not found)")
			Exit Sub
		End If
		
		CurrPos = 5
		
		'Extract the checksum
		CheckSum = ByteArray(CurrPos - 1)
		CurrPos = CurrPos + 1
		
		'Extract the length of the original string
		Call CopyMem(ResultLen, ByteArray(CurrPos - 1), 4)
		CurrPos = CurrPos + 4
		lResultLen = ResultLen
		
		'If the compressed string is empty we can
		'skip the function right here
		If (ResultLen = 0) Then Exit Sub
		
		'Create the result array
		ReDim Result(ResultLen - 1)
		
		'Get the number of characters used
		Call CopyMem(Count, ByteArray(CurrPos - 1), 2)
		CurrPos = CurrPos + 2
		
		'Get the used characters and their
		'respective bit sequence lengths
		For i = 1 To Count
			With CharValue(ByteArray(CurrPos - 1))
				CurrPos = CurrPos + 1
				.Count = ByteArray(CurrPos - 1)
				CurrPos = CurrPos + 1
				ReDim .Data(.Count - 1)
			End With
		Next 
		
		'Create a small array to hold the bit values,
		'this is (still) faster than calculating on-fly
		For i = 0 To 7
			BitValue(i) = 2 ^ i
		Next 
		
		'Extract the Huffman Tree, converting the
		'byte sequence to bit sequences
		ByteValue = ByteArray(CurrPos - 1)
		CurrPos = CurrPos + 1
		BitPos = 0
		For i = 0 To 255
			With CharValue(i)
				If (.Count > 0) Then
					For j = 0 To (.Count - 1)
						If (ByteValue And BitValue(BitPos)) Then .Data(j) = 1
						BitPos = BitPos + 1
						If (BitPos = 8) Then
							ByteValue = ByteArray(CurrPos - 1)
							CurrPos = CurrPos + 1
							BitPos = 0
						End If
					Next 
				End If
			End With
		Next 
		If (BitPos = 0) Then CurrPos = CurrPos - 1
		
		'Create the Huffman Tree
		NodesCount = 1
		Nodes(0).LeftNode = -1
		Nodes(0).RightNode = -1
		Nodes(0).ParentNode = -1
		Nodes(0).Value = -1
		For i = 0 To 255
			Call CreateTree(Nodes, NodesCount, i, CharValue(i))
		Next 
		
		'Decode the actual data
		ResultLen = 0
		For CurrPos = CurrPos To ByteLen
			ByteValue = ByteArray(CurrPos - 1)
			For BitPos = 0 To 7
				If (ByteValue And BitValue(BitPos)) Then
					NodeIndex = Nodes(NodeIndex).RightNode
				Else
					NodeIndex = Nodes(NodeIndex).LeftNode
				End If
				If (Nodes(NodeIndex).Value > -1) Then
					Result(ResultLen) = Nodes(NodeIndex).Value
					ResultLen = ResultLen + 1
					If (ResultLen = lResultLen) Then GoTo DecodeFinished
					NodeIndex = 0
				End If
			Next 
			If (CurrPos Mod 10000 = 0) Then
				NewProgress = CurrPos / ByteLen * PROGRESS_DECODING
				If (NewProgress <> CurrProgress) Then
					CurrProgress = NewProgress
					RaiseEvent Progress(CurrProgress)
				End If
			End If
		Next 
DecodeFinished: 
		
		'Verify data to check for corruption.
		Char_Renamed = 0
		For i = 0 To (ResultLen - 1)
			Char_Renamed = Char_Renamed Xor Result(i)
			If (i Mod 10000 = 0) Then
				NewProgress = i / ResultLen * PROGRESS_CHECKCRC + PROGRESS_DECODING
				If (NewProgress <> CurrProgress) Then
					CurrProgress = NewProgress
					RaiseEvent Progress(CurrProgress)
				End If
			End If
		Next 
		If (Char_Renamed <> CheckSum) Then
			Err.Raise(vbObjectError, "clsHuffman.Decode()", "The data might be corrupted (checksum did not match expected value)")
		End If
		
		'Return the uncompressed string
		ReDim ByteArray(ResultLen - 1)
		Call CopyMem(ByteArray(0), Result(0), ResultLen)
		
		'Make sure we get a "100%" progress message
		If (CurrProgress <> 100) Then
			RaiseEvent Progress(100)
		End If
		
	End Sub
	
	
	Private Function FileExist(ByRef Filename As String) As Boolean
		
		On Error GoTo FileDoesNotExist
		
		Call FileLen(Filename)
		FileExist = True
		Exit Function
		
FileDoesNotExist: 
		FileExist = False
		
	End Function
End Class