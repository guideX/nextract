Option Strict Off
Option Explicit On
Module mdlLZSSDeCompress
	'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	'++ I adapted this code from Marco v/d Berg's Compression Methods+_ V 1.04 code           ++
	'++ http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=37867&lngWId=1    ++
	'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	'++ I removed parts uneeded and made very few additions... the code is pretty much intact ++
	'++ The LZSS routine can be brutally slow on the input side!... but is faster on the      ++
	'++ output side due (mostly) to the lack of string handling.  Good work Marco v/d Berg!   ++
	'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	'This is a 1 run method but we have to keep the whole contents
	'in memory until some variables are saved wich are needed bij the decompressor
	'This LZSS routine make its compares in strings to find matches
	'During DeCompression no strings are used just the bytestream that has past
	
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Public Declare Sub CopyMem Lib "kernel32"  Alias "RtlMoveMemory"(ByRef Destination As Any, ByRef Source As Any, ByVal Length As Integer)
	Public OriginalArray() As Byte 'array to store the original
	'Public OriginalSize As Long             'size of the original file
	Public Const DictionarySize As Short = 16 'larger seems to produce better compression... but slower
	
	
	Private Structure LZSSStream
		Dim Data() As Byte
		Dim Position As Integer
		Dim BitPos As Byte
		Dim Buffer As Byte
	End Structure
	Private Stream(3) As LZSSStream '0=controlstream   1=distenceStream  2=lengthstream   3=literalstream
	Private MaxHistory As Integer
	Public Sub Decompress_LZSS3(ByRef ByteArray() As Byte)
		
		Dim X As Integer
		Dim InPos As Integer
		Dim Temp As Integer
		Dim ContPos As Integer
		Dim ContBit As Byte
		Dim DistPos As Integer
		Dim LengthPos As Integer
		Dim LitPos As Integer
		Dim Data As Short
		Dim Distance As Integer
		Dim Length As Short
		Dim CopyPos As Integer
		Dim AddText As String
		ReDim Stream(0).Data(500)
		Stream(0).BitPos = 0
		Stream(0).Buffer = 0
		Stream(0).Position = 0
		ContPos = 6
		ContBit = 0
		Temp = CInt(ByteArray(0)) * 256 + ByteArray(1)
		Temp = CInt(Temp) * 256 + ByteArray(2)
		DistPos = ContPos + Temp + 1
		Temp = CInt(ByteArray(3)) * 256 + ByteArray(4)
		Temp = CInt(Temp) * 256 + ByteArray(5)
		LengthPos = Temp + Temp + DistPos + 2 + 2
		LitPos = LengthPos + Temp + 1
		MaxHistory = CInt(1024) * ByteArray(LitPos)
		LitPos = LitPos + 1
		Call AddBitsToStream(Stream(0), CInt(ByteArray(LitPos)), 8)
		LitPos = LitPos + 1
		Do 
			If ReadBitsFromArray(ByteArray, ContPos, ContBit, 1) = 0 Then
				'read literal data
				Call AddBitsToStream(Stream(0), ReadBitsFromArray(ByteArray, LitPos, 0, 8), 8)
			Else
				Distance = ReadBitsFromArray(ByteArray, DistPos, 0, 8)
				Distance = CInt(Distance) * 256 + ReadBitsFromArray(ByteArray, DistPos, 0, 8)
				If Distance = 0 Then
					Exit Do
				End If
				Length = ReadBitsFromArray(ByteArray, LengthPos, 0, 8) + 3
				CopyPos = Stream(0).Position - Distance
				For X = 0 To Length - 1
					Call AddBitsToStream(Stream(0), CByte(Stream(0).Data(CopyPos + X)), 8)
				Next 
			End If
		Loop 
		ReDim ByteArray(Stream(0).Position - 1)
		For X = 0 To Stream(0).Position - 1
			ByteArray(X) = Stream(0).Data(X)
		Next 
		
	End Sub
	'this sub will add an amount of bits to a certain stream
	Private Sub AddBitsToStream(ByRef Toarray As LZSSStream, ByRef Number As Byte, ByRef Numbits As Byte)
		Dim X As Integer
		If Numbits = 8 And Toarray.BitPos = 0 Then
			If Toarray.Position > UBound(Toarray.Data) Then ReDim Preserve Toarray.Data(Toarray.Position + 500)
			Toarray.Data(Toarray.Position) = Number And &HFFs
			Toarray.Position = Toarray.Position + 1
			Exit Sub
		End If
		For X = Numbits - 1 To 0 Step -1
			Toarray.Buffer = Toarray.Buffer * 2 + (-1 * CShort((Number And 2 ^ X) > 0))
			Toarray.BitPos = Toarray.BitPos + 1
			If Toarray.BitPos = 8 Then
				If Toarray.Position > UBound(Toarray.Data) Then ReDim Preserve Toarray.Data(Toarray.Position + 500)
				Toarray.Data(Toarray.Position) = Toarray.Buffer
				Toarray.BitPos = 0
				Toarray.Buffer = 0
				Toarray.Position = Toarray.Position + 1
			End If
		Next 
	End Sub
	'this sub will read an amount of bits from the inputstream
	Private Function ReadBitsFromArray(ByRef FromArray() As Byte, ByRef FromPos As Integer, ByRef FromBit As Byte, ByRef Numbits As Short) As Integer
		Dim X As Short
		Dim Temp As Integer
		If FromBit = 0 And Numbits = 8 Then
			ReadBitsFromArray = FromArray(FromPos)
			FromPos = FromPos + 1
			Exit Function
		End If
		For X = 1 To Numbits
			Temp = Temp * 2 + (-1 * CShort((FromArray(FromPos) And 2 ^ (7 - FromBit)) > 0))
			FromBit = FromBit + 1
			If FromBit = 8 Then
				If FromPos + 1 > UBound(FromArray) Then
					Do While X < Numbits
						Temp = Temp * 2
						X = X + 1
					Loop 
					FromPos = FromPos + 1
					Exit For
				End If
				FromPos = FromPos + 1
				FromBit = 0
			End If
		Next 
		ReadBitsFromArray = Temp
	End Function
End Module