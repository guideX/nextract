Option Strict Off
Option Explicit On
Module mdlProba
	Private sBoxInit(15) As Object
	Private BoxTurnOver As Object
	Private sBoxPos(32) As Byte
	Private sBox(32) As Byte
	Private sBoxOut(32) As Byte
	Private sBoxInvInit(32, 15) As Byte
	
	Public Function EncodeString(ByRef aText As String, ByVal aKey As String, ByRef expandKey As Boolean) As String
		On Error GoTo ErrHandler
		Dim i As Double
		If expandKey = False Then
			Call PROBAsetKey(aKey)
		Else
			Call PROBAsetExpandedKey(aKey)
		End If
		For i = 1 To Len(aText)
			EncodeString = EncodeString & Chr(PROBAencodeByte(Asc(Mid(aText, i, 1))))
		Next i
		Exit Function
ErrHandler: 
		'ProcessRuntimeError Err.Number, Err.Description, "Public Function EncodeString(aText As String, ByVal aKey As String, expandKey As Boolean) As String"
	End Function
	
	Public Function DecodeString(ByRef aText As String, ByVal aKey As String, ByRef expandKey As Boolean) As String
		On Error GoTo ErrHandler
		Dim i As Double
		If expandKey = False Then
			Call PROBAsetKey(aKey)
		Else
			Call PROBAsetExpandedKey(aKey)
		End If
		For i = 1 To Len(aText)
			DecodeString = DecodeString & Chr(PROBAdecodeByte(Asc(Mid(aText, i, 1))))
		Next i
		Exit Function
ErrHandler: 
		'ProcessRuntimeError Err.Number, Err.Description, "Public Function DecodeString(aText As String, ByVal aKey As String, expandKey As Boolean) As String"
	End Function
	
	Private Sub PROBAinitKey()
		On Error GoTo ErrHandler
		Dim i, b As Byte
		'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		'UPGRADE_WARNING: Couldn't resolve default property of object sBoxInit(0). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sBoxInit(0) = New Object(){12, 8, 9, 1, 2, 4, 10, 13, 11, 3, 0, 15, 7, 6, 14, 5}
		'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		'UPGRADE_WARNING: Couldn't resolve default property of object sBoxInit(1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sBoxInit(1) = New Object(){4, 0, 10, 11, 3, 14, 9, 8, 13, 1, 2, 7, 5, 6, 12, 15}
		'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		'UPGRADE_WARNING: Couldn't resolve default property of object sBoxInit(2). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sBoxInit(2) = New Object(){0, 12, 15, 2, 4, 3, 9, 13, 1, 10, 8, 11, 14, 5, 7, 6}
		'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		'UPGRADE_WARNING: Couldn't resolve default property of object sBoxInit(3). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sBoxInit(3) = New Object(){7, 6, 8, 5, 0, 9, 3, 2, 1, 10, 15, 11, 14, 4, 13, 12}
		'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		'UPGRADE_WARNING: Couldn't resolve default property of object sBoxInit(4). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sBoxInit(4) = New Object(){11, 13, 4, 3, 9, 10, 5, 1, 8, 12, 6, 14, 7, 15, 2, 0}
		'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		'UPGRADE_WARNING: Couldn't resolve default property of object sBoxInit(5). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sBoxInit(5) = New Object(){14, 3, 4, 1, 0, 10, 5, 11, 2, 15, 6, 8, 12, 13, 9, 7}
		'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		'UPGRADE_WARNING: Couldn't resolve default property of object sBoxInit(6). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sBoxInit(6) = New Object(){1, 5, 11, 12, 6, 4, 15, 0, 7, 3, 14, 9, 13, 8, 10, 2}
		'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		'UPGRADE_WARNING: Couldn't resolve default property of object sBoxInit(7). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sBoxInit(7) = New Object(){5, 3, 11, 13, 2, 1, 12, 10, 0, 4, 7, 6, 14, 8, 15, 9}
		'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		'UPGRADE_WARNING: Couldn't resolve default property of object sBoxInit(8). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sBoxInit(8) = New Object(){11, 3, 10, 5, 1, 14, 12, 13, 15, 2, 7, 8, 6, 0, 9, 4}
		'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		'UPGRADE_WARNING: Couldn't resolve default property of object sBoxInit(9). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sBoxInit(9) = New Object(){1, 2, 6, 0, 15, 5, 13, 3, 14, 4, 10, 12, 9, 11, 8, 7}
		'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		'UPGRADE_WARNING: Couldn't resolve default property of object sBoxInit(10). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sBoxInit(10) = New Object(){12, 11, 13, 3, 2, 14, 9, 4, 1, 10, 8, 7, 0, 6, 5, 15}
		'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		'UPGRADE_WARNING: Couldn't resolve default property of object sBoxInit(11). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sBoxInit(11) = New Object(){3, 10, 4, 5, 0, 9, 6, 8, 7, 11, 12, 13, 2, 15, 14, 1}
		'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		'UPGRADE_WARNING: Couldn't resolve default property of object sBoxInit(12). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sBoxInit(12) = New Object(){7, 0, 9, 8, 3, 10, 13, 1, 11, 4, 2, 12, 6, 14, 5, 15}
		'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		'UPGRADE_WARNING: Couldn't resolve default property of object sBoxInit(13). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sBoxInit(13) = New Object(){5, 11, 4, 3, 2, 0, 12, 1, 15, 14, 6, 10, 9, 13, 7, 8}
		'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		'UPGRADE_WARNING: Couldn't resolve default property of object sBoxInit(14). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sBoxInit(14) = New Object(){3, 2, 1, 10, 11, 9, 15, 4, 5, 14, 13, 0, 6, 7, 12, 8}
		'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		'UPGRADE_WARNING: Couldn't resolve default property of object sBoxInit(15). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sBoxInit(15) = New Object(){2, 6, 13, 0, 15, 14, 12, 9, 8, 11, 3, 10, 5, 7, 4, 1}
		For i = 0 To 15
			For b = 0 To 15
				'UPGRADE_WARNING: Couldn't resolve default property of object sBoxInit()(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sBoxInvInit(i, sBoxInit(i)(b)) = b
			Next 
		Next i
		'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		'UPGRADE_WARNING: Couldn't resolve default property of object BoxTurnOver. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BoxTurnOver = New Object(){0, 2, 14, 6, 15, 3, 7, 11, 5, 9, 3, 14, 13, 4, 6, 8, 1}
		Exit Sub
ErrHandler: 
		'ProcessRuntimeError Err.Number, Err.Description, "Private Sub PROBAinitKey()"
	End Sub
	
	Private Sub PROBAsetKey(ByRef aKey As String)
		On Error GoTo ErrHandler
		Dim i As Byte
		Dim n(32) As Byte
		Dim Key() As Byte
		'UPGRADE_ISSUE: Constant vbFromUnicode was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		'UPGRADE_TODO: Code was upgraded to use System.Text.UnicodeEncoding.Unicode.GetBytes() which may not have the same behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="93DD716C-10E3-41BE-A4A8-3BA40157905B"'
		Key = System.Text.UnicodeEncoding.Unicode.GetBytes(StrConv(aKey, vbFromUnicode))
		For i = 0 To 15
			sBox(i + 1) = Key(i) And 15
			sBox(i + 17) = Int(Key(i) / 16)
		Next 
		For i = 0 To 15
			sBoxPos(i + 1) = Key(i + 16) And 15
			sBoxPos(i + 17) = Int(Key(i + 16) / 16)
		Next 
		Call PROBAinitKey()
		Exit Sub
ErrHandler: 
		'ProcessRuntimeError Err.Number, Err.Description, "Private Sub PROBAsetKey(aKey As String)"
	End Sub
	
	Private Sub PROBAsetExpandedKey(ByRef aKey As String)
		On Error GoTo ErrHandler
		Dim g, t, i, e, r As Short
		Dim n(255) As Short
		Dim k() As Byte
		Dim b As Byte
		For i = 0 To 255
			n(i) = i
		Next 
		r = Len(aKey)
		'UPGRADE_ISSUE: Constant vbFromUnicode was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		'UPGRADE_TODO: Code was upgraded to use System.Text.UnicodeEncoding.Unicode.GetBytes() which may not have the same behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="93DD716C-10E3-41BE-A4A8-3BA40157905B"'
		k = System.Text.UnicodeEncoding.Unicode.GetBytes(StrConv(aKey, vbFromUnicode))
		For i = 0 To 255
			g = (g + n(i) + k(i Mod r)) Mod 255
			b = n(i)
			n(i) = n(g)
			n(g) = b
		Next 
		For i = 1 To 16
			sBox(i) = n(i) And 15
			sBox(i + 16) = Int(n(i) / 16)
		Next 
		For i = 1 To 16
			sBoxPos(i) = n(16 + i) And 15
			sBoxPos(i + 16) = Int(n(16 + i) / 16)
		Next 
		Call PROBAinitKey()
		Exit Sub
ErrHandler: 
		'ProcessRuntimeError Err.Number, Err.Description, "Private Sub PROBAsetExpandedKey(aKey As String)"
	End Sub
	
	Private Function PROBAencodeByte(ByRef aByte As Byte) As Byte
		On Error GoTo ErrHandler
		sBoxOut(1) = sBoxEncode(Int(aByte / 16), 1)
		sBoxOut(2) = sBoxEncode(sBoxOut(1), 2)
		sBoxOut(3) = sBoxEncode(sBoxOut(2), 3)
		sBoxOut(4) = sBoxEncode(sBoxOut(3), 4)
		sBoxOut(5) = sBoxEncode(sBoxOut(4), 5)
		sBoxOut(6) = sBoxEncode(sBoxOut(5), 6)
		sBoxOut(7) = sBoxEncode(sBoxOut(6), 7)
		sBoxOut(8) = sBoxEncode(sBoxOut(7), 8)
		sBoxOut(9) = sBoxEncode(sBoxOut(8), 9)
		sBoxOut(10) = sBoxEncode(sBoxOut(9), 10)
		sBoxOut(11) = sBoxEncode(sBoxOut(10), 11)
		sBoxOut(12) = sBoxEncode(sBoxOut(11), 12)
		sBoxOut(13) = sBoxEncode(sBoxOut(12), 13)
		sBoxOut(14) = sBoxEncode(sBoxOut(13), 14)
		sBoxOut(15) = sBoxEncode(sBoxOut(14), 15)
		sBoxOut(16) = sBoxEncode(sBoxOut(15), 16)
		sBoxOut(17) = sBoxEncode(aByte And 15, 17)
		sBoxOut(18) = sBoxEncode(sBoxOut(17), 18)
		sBoxOut(19) = sBoxEncode(sBoxOut(18), 19)
		sBoxOut(20) = sBoxEncode(sBoxOut(19), 20)
		sBoxOut(21) = sBoxEncode(sBoxOut(20), 21)
		sBoxOut(22) = sBoxEncode(sBoxOut(21), 22)
		sBoxOut(23) = sBoxEncode(sBoxOut(22), 23)
		sBoxOut(24) = sBoxEncode(sBoxOut(23), 24)
		sBoxOut(25) = sBoxEncode(sBoxOut(24), 25)
		sBoxOut(26) = sBoxEncode(sBoxOut(25), 26)
		sBoxOut(27) = sBoxEncode(sBoxOut(26), 27)
		sBoxOut(28) = sBoxEncode(sBoxOut(27), 28)
		sBoxOut(29) = sBoxEncode(sBoxOut(28), 29)
		sBoxOut(30) = sBoxEncode(sBoxOut(29), 30)
		sBoxOut(31) = sBoxEncode(sBoxOut(30), 31)
		sBoxOut(32) = sBoxEncode(sBoxOut(31), 32)
		PROBAencodeByte = (sBoxOut(16) * 16) + sBoxOut(32)
		Call PROBAturnBoxes()
		Exit Function
ErrHandler: 
		'ProcessRuntimeError Err.Number, Err.Description, "Private Function PROBAencodeByte(aByte As Byte) As Byte"
	End Function
	
	Private Function PROBAdecodeByte(ByRef aByte As Byte) As Byte
		On Error GoTo ErrHandler
		Dim b, y As Byte
		sBoxOut(16) = Int(aByte / 16)
		sBoxOut(15) = sBoxDecode(sBoxOut(16), 16)
		sBoxOut(14) = sBoxDecode(sBoxOut(15), 15)
		sBoxOut(13) = sBoxDecode(sBoxOut(14), 14)
		sBoxOut(12) = sBoxDecode(sBoxOut(13), 13)
		sBoxOut(11) = sBoxDecode(sBoxOut(12), 12)
		sBoxOut(10) = sBoxDecode(sBoxOut(11), 11)
		sBoxOut(9) = sBoxDecode(sBoxOut(10), 10)
		sBoxOut(8) = sBoxDecode(sBoxOut(9), 9)
		sBoxOut(7) = sBoxDecode(sBoxOut(8), 8)
		sBoxOut(6) = sBoxDecode(sBoxOut(7), 7)
		sBoxOut(5) = sBoxDecode(sBoxOut(6), 6)
		sBoxOut(4) = sBoxDecode(sBoxOut(5), 5)
		sBoxOut(3) = sBoxDecode(sBoxOut(4), 4)
		sBoxOut(2) = sBoxDecode(sBoxOut(3), 3)
		sBoxOut(1) = sBoxDecode(sBoxOut(2), 2)
		b = sBoxDecode(sBoxOut(1), 1)
		sBoxOut(32) = aByte And 15
		sBoxOut(31) = sBoxDecode(sBoxOut(32), 32)
		sBoxOut(30) = sBoxDecode(sBoxOut(31), 31)
		sBoxOut(29) = sBoxDecode(sBoxOut(30), 30)
		sBoxOut(28) = sBoxDecode(sBoxOut(29), 29)
		sBoxOut(27) = sBoxDecode(sBoxOut(28), 28)
		sBoxOut(26) = sBoxDecode(sBoxOut(27), 27)
		sBoxOut(25) = sBoxDecode(sBoxOut(26), 26)
		sBoxOut(24) = sBoxDecode(sBoxOut(25), 25)
		sBoxOut(23) = sBoxDecode(sBoxOut(24), 24)
		sBoxOut(22) = sBoxDecode(sBoxOut(23), 23)
		sBoxOut(21) = sBoxDecode(sBoxOut(22), 22)
		sBoxOut(20) = sBoxDecode(sBoxOut(21), 21)
		sBoxOut(19) = sBoxDecode(sBoxOut(20), 20)
		sBoxOut(18) = sBoxDecode(sBoxOut(19), 19)
		sBoxOut(17) = sBoxDecode(sBoxOut(18), 18)
		y = sBoxDecode(sBoxOut(17), 17)
		PROBAdecodeByte = (b * 16) + y
		Call PROBAturnBoxes()
		Exit Function
ErrHandler: 
		'ProcessRuntimeError Err.Number, Err.Description, "Private Function PROBAdecodeByte(aByte As Byte) As Byte"
	End Function
	
	Private Function sBoxEncode(ByRef aByte As Byte, ByRef aBox As Byte) As Byte
		On Error GoTo ErrHandler
		Dim b As Byte
		b = aByte + sBoxPos(aBox)
		If b > 15 Then b = b - 16
		'UPGRADE_WARNING: Couldn't resolve default property of object sBoxInit()(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sBoxEncode = sBoxInit(sBox(aBox))(b)
		Exit Function
ErrHandler: 
		'ProcessRuntimeError Err.Number, Err.Description, "Private Function sBoxEncode(aByte As Byte, aBox As Byte) As Byte"
	End Function
	
	Private Function sBoxDecode(ByRef aByte As Byte, ByRef aBox As Byte) As Byte
		On Error GoTo ErrHandler
		Dim i As Short
		i = sBoxInvInit(sBox(aBox), aByte)
		i = i - sBoxPos(aBox)
		If i < 0 Then i = i + 16
		sBoxDecode = i
		Exit Function
ErrHandler: 
		'ProcessRuntimeError Err.Number, Err.Description, "Private Function sBoxDecode(aByte As Byte, aBox As Byte) As Byte"
	End Function
	
	Private Sub PROBAturnBoxes()
		On Error GoTo ErrHandler
		Dim i As Byte
		Rotate((1))
		For i = 1 To 15
			'UPGRADE_WARNING: Couldn't resolve default property of object BoxTurnOver(sBox(i)). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If sBoxPos(i) = BoxTurnOver(sBox(i)) Then
				Rotate((i + 1))
			Else
				Exit For
			End If
		Next 
		Rotate(17)
		For i = 17 To 31
			'UPGRADE_WARNING: Couldn't resolve default property of object BoxTurnOver(sBox(i)). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If sBoxPos(i) = BoxTurnOver(sBox(i)) Then
				Rotate((i + 1))
			Else
				Exit For
			End If
		Next 
		If sBoxOut(1) = 0 Then Rotate((26))
		If sBoxOut(1) = 0 Then Rotate((23))
		If sBoxOut(17) = 0 Then Rotate((14))
		If sBoxOut(17) = 0 Then Rotate((8))
		If sBoxOut(3) = 0 Then Rotate((21))
		If sBoxOut(18) = 0 Then Rotate((6))
		If sBoxOut(2) = 0 And sBoxOut(4) = 0 Then Rotate((28))
		If sBoxOut(7) = 0 And sBoxOut(12) = 0 Then Rotate((15))
		If sBoxOut(20) = 0 And sBoxOut(24) = 0 Then Rotate((7))
		If sBoxOut(5) = 0 And sBoxOut(6) = 0 Then Rotate((31))
		If sBoxOut(18) = 0 And sBoxOut(20) = 0 Then Rotate((17))
		If sBoxOut(6) + sBoxOut(27) = 8 Then Rotate((25))
		If sBoxOut(10) + sBoxOut(19) = 8 Then Rotate((5))
		If sBoxOut(8) + sBoxOut(21) = 8 Then Rotate((30))
		If sBoxOut(7) + sBoxOut(19) = 8 Then Rotate((9))
		If sBoxOut(4) + sBoxOut(7) = 8 Then Rotate((10))
		If sBoxOut(2) + sBoxOut(19) = 15 Then Rotate((32))
		If sBoxOut(3) + sBoxOut(22) = 15 Then Rotate((16))
		If sBoxOut(6) + sBoxOut(21) = 15 Then Rotate((11))
		If sBoxOut(7) + sBoxOut(19) = 15 Then Rotate((19))
		Dim tmp As String
		For i = 1 To 16
			tmp = Trim(Str(sBoxPos(i))) & " "
			If Len(tmp) = 2 Then tmp = "0" & tmp
		Next 
		For i = 17 To 32
			tmp = Trim(Str(sBoxPos(i))) & " "
			If Len(tmp) = 2 Then tmp = "0" & tmp
		Next i
		Exit Sub
ErrHandler: 
		'ProcessRuntimeError Err.Number, Err.Description, "Private Sub PROBAturnBoxes()"
	End Sub
	
	Private Sub Rotate(ByRef aPos As Byte)
		On Error GoTo ErrHandler
		sBoxPos(aPos) = sBoxPos(aPos) + 1
		If sBoxPos(aPos) > 15 Then sBoxPos(aPos) = sBoxPos(aPos) - 16
		Exit Sub
ErrHandler: 
		'ProcessRuntimeError Err.Number, Err.Description, "Private Sub Rotate(aPos As Byte)"
	End Sub
End Module