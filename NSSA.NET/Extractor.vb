Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmMain
	Inherits System.Windows.Forms.Form
	Private lFrameIndex As Short
	Private WithEvents Huffman As clsHuffman
	Private MyExeName As String
	Private MyPath As String
	'UPGRADE_ISSUE: PropertyBag object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
	Private PropBag As New PropertyBag
	Private lAuthor As String
	Private lProjectName As String
	Private lLicenseAgreement As String
	Private lPassword As String
	Private lEncryptionPassword As String
	
	Private Sub StartExpanding()
		Dim pdone As Object
		Dim F As Object
		Dim newfile As Object
		Dim l As Object
		Dim X As Object
		On Error GoTo HndlError
		XP_ProgressBar1.Value = 0
		FileProgressLbl.Text = "0%" : System.Windows.Forms.Application.DoEvents()
		FileProgressLbl.Visible = True
		Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
		Dim FileCount As Integer
		Dim Filename As String
		Dim NewFileName As String
		Dim ByteArr() As Byte
		Dim mbox As MsgBoxResult
		'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FileCount = Val(PropBag.ReadProperty("FileCount"))
		TaskLbl.Visible = True
		TaskLbl.Text = "Initializing..." : System.Windows.Forms.Application.DoEvents()
		Dim ThisFile As String
		For X = 1 To FileCount
BackUp: 
			'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Filename = PropBag.ReadProperty("File" & X & "Name")
			'UPGRADE_WARNING: Couldn't resolve default property of object l. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			l = Len(Filename)
			'UPGRADE_WARNING: Couldn't resolve default property of object l. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			NewFileName = MyPath & Mid(Filename, 2, l - 1)
			If DoesFileExist(NewFileName) = True Then
				'UPGRADE_WARNING: Couldn't resolve default property of object newfile. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mbox = MsgBox("File already exists '" & newfile & "'. Would you like to replace this file?", MsgBoxStyle.YesNo + MsgBoxStyle.Question)
				If mbox = MsgBoxResult.No Then
					MsgBox("File '" & NewFileName & "' was not extracted", MsgBoxStyle.Critical)
					GoTo BackUp
				End If
			End If
			If VB.Right(LCase(Trim(NewFileName)), 4) = ".exe" Then lstFiles.Items.Add(NewFileName)
			ThisFile = MyPath & "NXTRCT0R.TMP"
			If VB.Left(Filename, 1) = "L" Then
				TaskLbl.Text = "Loading Compressed File..." : System.Windows.Forms.Application.DoEvents()
				'UPGRADE_WARNING: Couldn't resolve default property of object F. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				F = FreeFile
				'UPGRADE_WARNING: Couldn't resolve default property of object F. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				FileOpen(F, ThisFile, OpenMode.Output)
				'UPGRADE_WARNING: Couldn't resolve default property of object F. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				FileClose(F)
				'UPGRADE_WARNING: Couldn't resolve default property of object F. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				F = FreeFile
				'UPGRADE_WARNING: Couldn't resolve default property of object F. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				FileOpen(F, ThisFile, OpenMode.Binary)
				'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ByteArr = PropBag.ReadProperty("File" & X)
				'UPGRADE_WARNING: Couldn't resolve default property of object F. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				FilePut(F, ByteArr)
				'UPGRADE_WARNING: Couldn't resolve default property of object F. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				FileClose(F)
				'UPGRADE_WARNING: Couldn't resolve default property of object F. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				F = FreeFile
				'UPGRADE_WARNING: Couldn't resolve default property of object F. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				FileOpen(F, ThisFile, OpenMode.Binary)
				'UPGRADE_WARNING: Couldn't resolve default property of object F. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ReDim OriginalArray(LOF(F) - 1)
				'UPGRADE_WARNING: Couldn't resolve default property of object F. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				FileGet(F, OriginalArray)
				'UPGRADE_WARNING: Couldn't resolve default property of object F. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				FileClose(F)
				'UPGRADE_WARNING: Couldn't resolve default property of object l. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				TaskLbl.Text = "Decompressing " & Mid(Filename, 2, l - 1) & "..." : System.Windows.Forms.Application.DoEvents()
				Call Decompress_LZSS3(OriginalArray)
				'UPGRADE_WARNING: Couldn't resolve default property of object F. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				F = FreeFile
				'UPGRADE_WARNING: Couldn't resolve default property of object F. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				FileOpen(F, NewFileName, OpenMode.Output)
				'UPGRADE_WARNING: Couldn't resolve default property of object F. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				FileClose(F)
				'UPGRADE_WARNING: Couldn't resolve default property of object F. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				F = FreeFile
				'UPGRADE_WARNING: Couldn't resolve default property of object F. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				FileOpen(F, NewFileName, OpenMode.Binary)
				'UPGRADE_WARNING: Couldn't resolve default property of object F. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				FilePut(F, OriginalArray)
				'UPGRADE_WARNING: Couldn't resolve default property of object F. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				FileClose(F)
			End If
			If VB.Left(Filename, 1) = "H" Then
				TaskLbl.Text = "Copying Temporary File..." : System.Windows.Forms.Application.DoEvents()
				'UPGRADE_WARNING: Couldn't resolve default property of object F. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				F = FreeFile
				'UPGRADE_WARNING: Couldn't resolve default property of object F. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				FileOpen(F, MyPath & "NXTRCT0R.TMP", OpenMode.Output)
				'UPGRADE_WARNING: Couldn't resolve default property of object F. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				FileClose(F)
				'UPGRADE_WARNING: Couldn't resolve default property of object F. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				F = FreeFile
				'UPGRADE_WARNING: Couldn't resolve default property of object F. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				FileOpen(F, MyPath & "NXTRCT0R.TMP", OpenMode.Binary)
				'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ByteArr = PropBag.ReadProperty("File" & X)
				'UPGRADE_WARNING: Couldn't resolve default property of object F. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				FilePut(F, ByteArr)
				'UPGRADE_WARNING: Couldn't resolve default property of object F. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				FileClose(F)
				'UPGRADE_WARNING: Couldn't resolve default property of object l. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				TaskLbl.Text = "Decompressing " & Mid(Filename, 2, l - 1) & "..." : System.Windows.Forms.Application.DoEvents()
				Call Huffman.DecodeFile(ThisFile, NewFileName)
			End If
			If VB.Left(Filename, 1) = "O" Then
				'UPGRADE_WARNING: Couldn't resolve default property of object l. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				TaskLbl.Text = "Copying " & Mid(Filename, 2, l - 1) & "..." : System.Windows.Forms.Application.DoEvents()
				'UPGRADE_WARNING: Couldn't resolve default property of object F. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				F = FreeFile
				'UPGRADE_WARNING: Couldn't resolve default property of object F. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				FileOpen(F, NewFileName, OpenMode.Output)
				'UPGRADE_WARNING: Couldn't resolve default property of object F. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				FileClose(F)
				'UPGRADE_WARNING: Couldn't resolve default property of object F. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				F = FreeFile
				'UPGRADE_WARNING: Couldn't resolve default property of object F. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				FileOpen(F, NewFileName, OpenMode.Binary)
				'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ByteArr = PropBag.ReadProperty("File" & X)
				'UPGRADE_WARNING: Couldn't resolve default property of object F. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				FilePut(F, ByteArr)
				'UPGRADE_WARNING: Couldn't resolve default property of object F. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				FileClose(F)
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object pdone. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			pdone = System.Math.Round((X / FileCount) * 100)
			'UPGRADE_WARNING: Couldn't resolve default property of object pdone. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			FileProgressLbl.Text = pdone & "%" : System.Windows.Forms.Application.DoEvents()
			'UPGRADE_WARNING: Couldn't resolve default property of object pdone. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			XP_ProgressBar1.Value = CShort(pdone)
		Next 
		TaskLbl.Text = "Finished!"
		'UPGRADE_WARNING: Couldn't resolve default property of object F. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		F = FreeFile
		'UPGRADE_WARNING: Couldn't resolve default property of object F. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FileOpen(F, ThisFile, OpenMode.Output)
		'UPGRADE_WARNING: Couldn't resolve default property of object F. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FileClose(F)
		Kill(ThisFile)
		Me.Cursor = System.Windows.Forms.Cursors.Default
		FileClose()
		cmdNext_Click(cmdNext, Nothing)
		Exit Sub
HndlError: 
		FileClose()
		MsgBox("Error: " & Err.Description, MsgBoxStyle.Critical)
		End
	End Sub
	
	Private Sub cmdBack_Click(ByVal Sender As System.Object, ByVal e As System.EventArgs) Handles cmdBack.Click
		Dim i As Short
		If cmdBack.Enabled = True Then
			lFrameIndex = lFrameIndex - 1
			For i = 0 To fraSetup.Count - 1
				fraSetup(i).Visible = False
			Next i
			fraSetup(lFrameIndex).Visible = True
		End If
	End Sub
	
	Private Sub cmdChangeDir_Click(ByVal Sender As System.Object, ByVal e As System.EventArgs) Handles cmdChangeDir.Click
		On Error GoTo ErrH
		Dim sPath As String
		sPath = SelectFolder(Me, "Select folder")
		If Len(sPath) = 0 Then
			lblPath.Text = ""
			Exit Sub
		Else
			If VB.Right(sPath, 1) <> "\" Then sPath = sPath & "\"
			lblPath.Text = sPath
			MyPath = sPath
		End If
		Exit Sub
ErrH: 
		MsgBox(Err.Number & Chr(10) & Err.Description)
	End Sub
	
	Private Sub cmdExit_Click(ByVal Sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
		Me.Close()
	End Sub
	
	Private Sub cmdFinish_Click(ByVal Sender As System.Object, ByVal e As System.EventArgs) Handles cmdFinish.Click
		If cmdFinish.Enabled = False Then Exit Sub
		Me.Close()
	End Sub
	
	Private Sub cmdNext_Click(ByVal Sender As System.Object, ByVal e As System.EventArgs) Handles cmdNext.Click
		Dim i As Short
		If cmdNext.Enabled = True Then
			If lFrameIndex = fraSetup.Count - 1 Then Exit Sub
			For i = 0 To fraSetup.Count - 1
				fraSetup(i).Visible = False
			Next i
			lFrameIndex = lFrameIndex + 1
			fraSetup(lFrameIndex).Visible = True
			Select Case lFrameIndex
				Case 3
					StartExpanding()
					cmdNext.Enabled = False
					cmdBack.Enabled = False
			End Select
			cmdBack.Enabled = True
			If lFrameIndex = (fraSetup.Count - 1) Then
				cmdNext.Enabled = False
				cmdFinish.Enabled = True
				cmdBack.Enabled = False
			End If
		End If
	End Sub
	
	Private Sub cmdRunFile_Click(ByVal Sender As System.Object, ByVal e As System.EventArgs) Handles cmdRunFile.Click
		On Error Resume Next
		Shell(lstFiles.Text)
	End Sub
	
	Public Function DoesFileExist(ByRef lFilename As String) As Boolean
		'On Local Error Resume Next
		Dim Msg As String
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		Msg = Dir(lFilename)
		If Msg <> "" Then
			DoesFileExist = True
		Else
			DoesFileExist = False
		End If
	End Function
	
	Public Function ReturnPassword() As String
		ReturnPassword = lPassword
	End Function
	
	Private Sub frmMain_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim F As Object
		On Error GoTo ErrHandler
		Huffman = New clsHuffman
		Dim BeginPos As Integer
		Dim tmpByte As Object
		Dim ByteArr() As Byte
		lEncryptionPassword = "820huiewfkj4hgru32o0ewyuwe"
		Image1.Image = Me.Icon.ToBitmap
		Image2.Image = Me.Icon.ToBitmap
		Image3.Image = Me.Icon.ToBitmap
		Image4.Image = Me.Icon.ToBitmap
		Image5.Image = Me.Icon.ToBitmap
		MyPath = My.Application.Info.DirectoryPath
		If VB.Right(MyPath, 1) <> "\" Then MyPath = MyPath & "\"
		'UPGRADE_WARNING: App property App.EXEName has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		MyExeName = MyPath & My.Application.Info.AssemblyName & ".exe"
		lblInformation.Text = "Ready to extract to:"
		lblPath.Text = MyPath
		'UPGRADE_WARNING: Couldn't resolve default property of object F. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		F = FreeFile
		'UPGRADE_WARNING: Couldn't resolve default property of object F. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FileOpen(F, MyExeName, OpenMode.Binary)
		'UPGRADE_WARNING: Couldn't resolve default property of object F. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(F, BeginPos, LOF(F) - 3)
		'UPGRADE_WARNING: Couldn't resolve default property of object F. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Seek(F, BeginPos)
		'UPGRADE_WARNING: Couldn't resolve default property of object F. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Get was upgraded to FileGetObject and has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGetObject(F, tmpByte)
		'UPGRADE_WARNING: Couldn't resolve default property of object F. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FileClose(F)
		'UPGRADE_WARNING: Couldn't resolve default property of object tmpByte. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ByteArr = tmpByte
		'UPGRADE_ISSUE: PropertyBag property PropBag.Contents was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		PropBag.Contents = ByteArr
		Me.Text = "Self Extracting Archive"
		'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		lProjectName = PropBag.ReadProperty("ProjectName", "")
		'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		lAuthor = PropBag.ReadProperty("Author", "")
		'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		txtLicenseAgreement.Text = PropBag.ReadProperty("LicenseAgreement", "")
		lblReady.Text = "Enough information has been collected about the target system to extract '" & lProjectName & "'. Click 'Next' to extract file(s)"
		lblYouRan.Text = "You have run the self extracting archive '" & lProjectName & "', click next to continue"
		'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		lPassword = PropBag.ReadProperty("Password", "")
		If Len(Trim(lPassword)) <> 0 Then
			lPassword = DecodeString(DecodeStr64(lPassword), lEncryptionPassword, True)
			If Len(lPassword) <> 0 Then
				frmPassword.ShowDialog()
			End If
		End If
		Me.Text = lProjectName & " by " & lAuthor & " - Self Extracting Archive"
		If Len(Trim(txtLicenseAgreement.Text)) = 0 Or Len(Trim(txtLicenseAgreement.Text)) = 1 Or Len(Trim(txtLicenseAgreement.Text)) = 2 Then
			fraLicenseAgreement.Visible = False
			cmdNext.Enabled = True
		End If
		Me.Visible = True
		Exit Sub
ErrHandler: 
		MsgBox(Err.Description)
		Me.Close()
	End Sub
	
	Private Sub isButton1_Click()
		FileClose()
		End
	End Sub
	
	Private Sub isButton4_Click()
		FileClose()
		End
	End Sub
	
	'UPGRADE_WARNING: Event optDisagree.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optDisagree_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optDisagree.CheckedChanged
		If eventSender.Checked Then
			cmdNext.Enabled = False
		End If
	End Sub
	
	'UPGRADE_WARNING: Event optIAgree.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optIAgree_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optIAgree.CheckedChanged
		If eventSender.Checked Then
			cmdNext.Enabled = True
		End If
	End Sub
End Class