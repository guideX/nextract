Option Strict Off
Option Explicit On
Module mdlChangeDir
	
	Public Structure BROWSEINFO
		Dim hOwner As Integer 'Handle to window's owner
		Dim pidlRoot As Integer 'Pointer to an item identifier list
		Dim pszDisplayName As String 'Pointer to a buffer that receives the display name of the folder selected
		Dim lpszTitle As String 'Pointer to a null-terminated string that is displayed above the tree view control in the dialog box
		Dim ulFlags As Integer 'Value specifying the types of folders to be listed in the dialog box as well as other options
		Dim lpfn As Integer 'Address an application-defined function that the dialog box calls when events occur
		Dim lParam As Integer 'Application-defined value that the dialog box passes to the callback function (if one is specified).
		Dim iImage As Integer 'Variable that receives the image associated with the selected folder. The image is specified as an index to the system image list.
	End Structure
	
	
	
	'***************************************************************
	'   Browse Dialog Flags & Constants
	'***************************************************************
	Public Const BIF_RETURNONLYFSDIRS As Short = &H1s 'Only returns file system directories
	Public Const BIF_DONTGOBELOWDOMAIN As Short = &H2s 'Does not include network folders below the domain level
	Public Const BIF_STATUSTEXT As Short = &H4s 'Includes a status area in the dialog box. The callback function can set the status text by sending messages to the dialog box.
	Public Const BIF_RETURNFSANCESTORS As Short = &H8s 'Only returns file system ancestors
	Public Const BIF_BROWSEFORCOMPUTER As Short = &H1000s 'Only returns computers
	Public Const BIF_BROWSEFORPRINTER As Short = &H2000s 'Only returns (network) printers
	
	Public Const MAX_PATH As Short = 255
	
	
	'***************************************************************
	'   Browse Dialog API Declarations
	'***************************************************************
	Public Declare Function SHGetPathFromIDList Lib "shell32.dll"  Alias "SHGetPathFromIDListA"(ByVal pidl As Integer, ByVal pszPath As String) As Integer
	
	'UPGRADE_WARNING: Structure BROWSEINFO may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Public Declare Function SHBrowseForFolder Lib "shell32.dll"  Alias "SHBrowseForFolderA"(ByRef lpBrowseInfo As BROWSEINFO) As Integer
	
	'UPGRADE_NOTE: pv was upgraded to pv_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv_Renamed As Integer)
	
	
	'***************************************************************
	'   opens the Browse Folder window and returns the folder selected
	'   as a string, or an empty string if canceled
	'***************************************************************
	Public Function SelectFolder(ByRef frm As System.Windows.Forms.Form, Optional ByRef sDialTitle As String = "Select a folder") As String
		
		Dim bi As BROWSEINFO
		Dim pidl As Integer
		Dim path As String
		Dim POS As Short
		
		'Fill the BROWSEINFO structure with the needed data.
		With bi
			
			.hOwner = frm.Handle.ToInt32
			.pidlRoot = 0 'Root folder to browse from, or desktop if Null
			.lpszTitle = sDialTitle 'Message to display in dialog
			.ulFlags = BIF_RETURNONLYFSDIRS 'the type of folder to return
			
		End With
		
		'show the browse for folders dialog
		pidl = SHBrowseForFolder(bi)
		
		'the dialog has closed, so parse & display the user's
		'returned folder selection contained in pidl
		path = Space(MAX_PATH)
		
		If SHGetPathFromIDList(pidl, path) Then
			POS = InStr(path, Chr(0))
			SelectFolder = Left(path, POS - 1)
		Else
			SelectFolder = ""
		End If
		
		Call CoTaskMemFree(pidl)
		
	End Function
End Module