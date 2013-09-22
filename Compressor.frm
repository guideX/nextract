VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{EE128208-4F73-11D3-83BB-C47C02EE3D01}#1.0#0"; "ControlResizer.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Nexgen Self Extract"
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8655
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Compressor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5730
   ScaleWidth      =   8655
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkCompressContents 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Use Compression"
      Height          =   255
      Left            =   5160
      TabIndex        =   28
      Top             =   4800
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Del"
      Height          =   375
      Left            =   7800
      TabIndex        =   27
      Top             =   3000
      Width           =   735
   End
   Begin NSEAM.XP_ProgressBar prg1 
      Height          =   375
      Left            =   1680
      TabIndex        =   25
      Top             =   3000
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BrushStyle      =   0
      Color           =   49344
      Scrolling       =   1
   End
   Begin VB.CommandButton Command5 
      Caption         =   "B&rowse"
      Height          =   285
      Left            =   7800
      TabIndex        =   8
      Top             =   3840
      Width           =   735
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4320
      Top             =   5280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Compressor.frx":29C12
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Compressor.frx":2A1AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Compressor.frx":2A74A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Compressor.frx":2ACE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Compressor.frx":2B282
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2775
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   4895
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "File"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Size"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "New Size"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Difference"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Filename"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Add"
      Height          =   375
      Left            =   6960
      TabIndex        =   6
      Top             =   3000
      Width           =   735
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Height          =   2295
      Left            =   120
      TabIndex        =   5
      Top             =   3360
      Width           =   4900
      Begin VB.Label FileProgressLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "0%"
         Height          =   255
         Left            =   1800
         TabIndex        =   24
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "File Progress:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label ArchSizeLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   1800
         TabIndex        =   22
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Archive Size:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label RemFileLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   1800
         TabIndex        =   20
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Remaining:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label TtlFileLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   1800
         TabIndex        =   18
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Files To Add:"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   1440
         Width           =   1635
      End
      Begin VB.Label TaskLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "None"
         Height          =   255
         Left            =   1800
         TabIndex        =   16
         Top             =   1200
         Width           =   3015
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Current Task:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label ComprSizeLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   1800
         TabIndex        =   14
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Compressed Size:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label SrcSizeLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   1800
         TabIndex        =   12
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Source Size:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label SrcFileNameLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "None"
         Height          =   255
         Left            =   1800
         TabIndex        =   10
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Source File:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   5160
      TabIndex        =   2
      Top             =   4440
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   5160
      TabIndex        =   1
      Top             =   3840
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Make Archive"
      Default         =   -1  'True
      Height          =   375
      Left            =   7080
      TabIndex        =   0
      Top             =   5280
      Width           =   1455
   End
   Begin ControlResizer.AutoResizer AutoResizer1 
      Height          =   615
      Left            =   6240
      Top             =   5160
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1085
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "File Progress:"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      X1              =   8520
      X2              =   4560
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "SelfExtractor FileName:"
      Height          =   315
      Index           =   1
      Left            =   5160
      TabIndex        =   4
      Top             =   4200
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Save in Folder:"
      Height          =   315
      Index           =   0
      Left            =   5160
      TabIndex        =   3
      Top             =   3600
      Width           =   3360
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit'... you that like option explicit can un-remark and chase down errors

Private WithEvents Huffman As clsHuffman
Attribute Huffman.VB_VarHelpID = -1
'**************************************
'Windows API/Global Declarations for :Common Dialog without OCX
'**************************************
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private strfileName As OPENFILENAME
Private FileSelected As Boolean '--> CptnVic's Addition... coerces value to boolean for ease of use... see form code for use.

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
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Dim TtlFileSize As Currency, TtlFiles As Integer, MyFileTitle As String, CFileName As String
Dim OutPutFileName As String, TmpFileHuff As String, TmpFileLZSS As String, TemplateFileName As String
Dim PropBag As New PropertyBag 'Make the propertybag, a usefull class
Dim ByteArr() As Byte 'Dim the byte array we use to put the files in and then in the propertybag
Dim FirstTime As Boolean
Private Sub DialogFilter(WantedFilter As String)
    Dim intLoopCount As Integer
    strfileName.lpstrFilter = ""
    For intLoopCount = 1 To Len(WantedFilter)
        If Mid(WantedFilter, intLoopCount, 1) = "|" Then strfileName.lpstrFilter = _
        strfileName.lpstrFilter + Chr(0) Else strfileName.lpstrFilter = _
        strfileName.lpstrFilter + Mid(WantedFilter, intLoopCount, 1)
    Next intLoopCount
    strfileName.lpstrFilter = strfileName.lpstrFilter + Chr(0)
End Sub

'This is The Function To get the File Name to Open
'Even If U don't specify a Title or a Filter it is OK
Private Function fncGetFileNametoOpen(Optional strDialogTitle As String = "Open", Optional strFilter As String = "All Files|*.*", Optional strDefaultExtention As String = "*.*") As String
    Dim lngReturnValue As Long
    Dim intRest As Integer
    strfileName.lpstrTitle = strDialogTitle
    strfileName.lpstrDefExt = strDefaultExtention
    DialogFilter (strFilter)
    strfileName.hInstance = App.hInstance
    strfileName.lpstrFile = Chr(0) & Space(259) ' --> will return Chr(0) & 259 spaces UNLESS a valid file is selected.
    strfileName.nMaxFile = 260 ' maximum length of a file name
    strfileName.flags = &H4
    strfileName.lStructSize = Len(strfileName)
    lngReturnValue = GetOpenFileName(strfileName)
    FileSelected = lngReturnValue ' --> CptnVics addition... must be done after the call to GetOpenFileName(strfileName)!
        'FileSelected will coerce this value (lngReturnValue) to boolean... true if a file was selected... false otherwise.
        'FileSelected could be dimensioned as a string... in which case it would return "1" if a file was selected... "0" if canceled
        'The boolean check takes less code.
    MyFileTitle = ""
    MyFileTitle = strfileName.lpstrTitle
    fncGetFileNametoOpen = strfileName.lpstrFile
End Function


'Private Sub Check1_Click()
'    If Check1.Value Then
'        Check1.Caption = "ON"
'    Else
'        Check1.Caption = "OFF"
'    End If
'End Sub

Private Sub Command1_Click()
'Stop
'MsgBox TemplateFileName
On Error GoTo ErrorHandler
    'build the archive
    'this error checking should be more robust... but if you want more... make some more
    If FirstTime = False Then
        Msg = "Re-start the compiler... or bad things may happen!"
        Title = "Re-Start 1st!"
        MsgBox Msg, vbOKOnly + vbExclamation, Title
        Exit Sub
    Else
        FirstTime = False
    End If
    'make sure installer template has been selected
    If TemplateFileName = "" Then
        MsgBox "You need to Set The Installer (template) File First!", vbOKOnly + vbExclamation
        Exit Sub
    End If
    'make sure save path is set
    If Len(Text1(0).Text) = 0 Then
        MsgBox "You need to Set The Output Path First!", vbOKOnly + vbExclamation
        Exit Sub
    End If
    If Right(UCase(Text1(0).Text), 1) <> "\" Then
        MsgBox "You need to Set The Output Path First!", vbOKOnly + vbExclamation
        Exit Sub
    End If
    'make sure filename to build exists
    If Len(Text1(1).Text) = 0 Then
        MsgBox "You need to Set The Output File Name First!", vbOKOnly + vbExclamation
        Exit Sub
    End If
    'make sure there are files to add
    If TtlFiles = 0 Then
        MsgBox "You need to Select Some Files To Add First!", vbOKOnly + vbExclamation
        Exit Sub
    End If
    'hopefully you haven't tried to outsmart the simple error checks!
    
    PropBag.WriteProperty "FileCount", TtlFiles 'Notify the self extracter how many files we have.
    
    'set the output and temp file name
    OutPutFileName = Text1(0).Text & Text1(1).Text
    TmpFileHuff = Text1(0).Text & "compareHuff.tmp"
    TmpFileLZSS = Text1(0).Text & "compareLZSS.tmp"
    'make sure if a self extractor exists... that it is destroyed 1st.
    F = FreeFile
        Open OutPutFileName For Output As #F
        Close #F
    'copy the self extractor exe to pile other files on
    FileCopy TemplateFileName, OutPutFileName
    
    'do the loop... compress... compare... store smaller of the three
    Dim Origsize As Long, HuffSize As Long, LZSize As Long, x As Integer, Smallest As Long
    'show user something is happening...
    Me.MousePointer = 11
    ToDo = Val(TtlFileLbl.Caption)
    RemFileLbl.Caption = ToDo
    For x = 1 To TtlFiles
        
        'Update the progress frame
        SrcFileNameLbl.Caption = ListView1.ListItems(x).ListSubItems(4).Text 'just the file name... no path
        SrcSizeLbl.Caption = ListView1.ListItems(x).ListSubItems(1).Text 'file with path intact
        ComprSizeLbl.Caption = "0"
        FileProgressLbl.Caption = "0%"
        ListView1.ListItems(x).SmallIcon = 1 'show user which file is being done
        CFileName = ListView1.ListItems(x).Text 'get name of file to compress
        Origsize = Val(ListView1.ListItems(x).ListSubItems(1).Text) 'store original file size
        Smallest = Origsize
        '-------------------------------------------
        'If chkCompressContents.Value = 0 Then
        '    bypass all compression
        '    TaskLbl.Caption = "Storing Original File"
        '    DoEvents
        '    StoreOriginal x, CFileName, Origsize
        '    ListView1.ListItems(x).ListSubItems(3).Text = 0
        '    GoTo SkipIt
        'End If
        'If chkCompressContents.Value = 0 Then GoTo SkipHuffman
        '    GoTo SkipHuffman
        'End If
        TaskLbl.Caption = "Huffman Compressing: " & SrcFileNameLbl.Caption
        ComprSizeLbl.Caption = "0"
        FileProgressLbl.Caption = "0%"
        DoEvents
        F = FreeFile
        Open TmpFileHuff For Output As #F 'clear existing file?
        Close #F
        HuffSize = 0
        'Huffman Compress the source file
        Call Huffman.EncodeFile(CFileName, TmpFileHuff)
        'get the compressed file size
        HuffSize = FileLen(TmpFileHuff)
        ComprSizeLbl.Caption = HuffSize
        'compare filesizes thus far
        If Origsize <= HuffSize Then
            Smallest = Origsize
        Else
            Smallest = HuffSize
        End If
        '-------------------------------------------
SkipHuffman:
        'do LZSS compression
        TaskLbl.Caption = "LZSS Compressing: " & SrcFileNameLbl.Caption
        DoEvents
        F = FreeFile
        Open CFileName For Binary As #F
            ReDim OriginalArray(0 To LOF(F) - 1)
            Get #F, , OriginalArray()
        Close #F
        'do lzss compression
        Call Compress_LZSS3(OriginalArray)
        'save file temporarily
        F = FreeFile
        Open TmpFileLZSS For Output As #F 'clear existing file?
        Close #F
        F = FreeFile
        Open TmpFileLZSS For Binary As #F
            Put #F, , OriginalArray()
        Close #F
        LZSize = FileLen(TmpFileLZSS)
        ComprSizeLbl.Caption = LZSize
        'compare filesizes thus far
        If Smallest <= LZSize Then
            Smallest = Origsize
        Else
            Smallest = LZSize
        End If
        '-------------------------------------------
        TaskLbl.Caption = "Storing File..."
        DoEvents
        '-------------------------------------------
        'compare best storage/retrieval method
        If Smallest = Origsize Then
            StoreOriginal x, CFileName, Origsize 'store copy of original
            ListView1.ListItems(x).ListSubItems(3).Text = 0
            GoTo SkipIt
        End If
        If Smallest = HuffSize Then
            StoreHuffman x, TmpFileHuff, HuffSize 'store huffman compressed version
            GoTo SkipIt
        End If
        If Smallest = LZSize Then
            StoreLZSS x, TmpFileLZSS, LZSize
        End If
        'calc savings if any
        Savings = 0
        Savings = Origsize - Smallest
        ListView1.ListItems(x).ListSubItems(3).Text = Savings
SkipIt:
        
        ToDo = ToDo - 1
        RemFileLbl.Caption = ToDo
        DoEvents
    Next
    TaskLbl.Caption = "Finished!"
    'the propertybag should have all it needs in it... so pile the files on the exe
    'Now we need to put the propertybag in the template exe:
    F = FreeFile
    Open OutPutFileName For Binary As #F 'Open the copied exe, in binary mode.
        Dim BeginPos As Long 'Dim variable to use for beginpos
        BeginPos = LOF(F) 'Get the total length of the file
        Seek #F, LOF(F) 'Start pointer at end of the file
        Put #F, , PropBag.Contents 'Put the whole propertybag with all the files in the template exe, so easy!
        Put #F, , BeginPos 'Add the long where the propertybag starts with it's data, so the self extracter knows it.
    Close #F 'Close the file.
    'clear up the temp files
    F = FreeFile
    Open TmpFileHuff For Output As #F
    Close #F
    Kill TmpFileHuff
    F = FreeFile
    Open TmpFileLZSS For Output As #F 'clear existing file?
    Close #F
    Kill TmpFileLZSS
    UpDateLast 'calc the total size expected
  

  Me.MousePointer = 0

  'Show a nice dialog to the user
  Savings = 0
  For x = 1 To TtlFiles
    Savings = Savings + Val(ListView1.ListItems(x).ListSubItems(3).Text)
  Next
  Outputs = FileLen(OutPutFileName)
  Msg = "Archive Complete" & vbCrLf & vbCrLf & "New Size: " & Outputs & vbCrLf & "Difference: " & Savings
  Call MsgBox(Msg, vbInformation + vbOKOnly, App.Title)
  Exit Sub
  
ErrorHandler:
  Close 'just in case
  Title = "An Error Occured During Installation"
  Msg = "Error # " & Str(Err.Number) & " was generated by " & Err.Source & Chr(13) & Err.Description
  MsgBox Msg, vbOKOnly + vbExclamation, Title
  End

End Sub
Private Sub StoreLZSS(x As Integer, TmpFileLZSS As String, LZSize As Long)
    'store the compressed version
    ListView1.ListItems(x).SmallIcon = 2 'show user which storage method selected
    ListView1.ListItems(x).ListSubItems(2).Text = LZSize
    ValNow = Val(ArchSizeLbl.Caption)
    ValNew = ValNow + LZSize
    ArchSizeLbl.Caption = ValNew
    TaskLbl.Caption = "Storing LZSS Compressed File"
    DoEvents
    F = FreeFile
    Open TmpFileLZSS For Binary As #F 'Open the file, in binary
        ReDim ByteArr(0 To LOF(F) - 1) 'clear the bytearray
        Get #F, , ByteArr() 'Get the file in our memory
    Close #F 'Close the file

    PropBag.WriteProperty "File" & x, ByteArr() 'Put the file in the propertybag, to get it back in the self extract exe.
    PropBag.WriteProperty "File" & x & "Name", "L" & SrcFileNameLbl.Caption 'Put the original filename, so the self extracter knows how to name the filename.
End Sub
Private Sub StoreHuffman(x As Integer, TmpFileHuff As String, HuffSize As Long)
    'store the compressed version
    ListView1.ListItems(x).SmallIcon = 3 'show user which storage method selected
    ListView1.ListItems(x).ListSubItems(2).Text = HuffSize
    ValNow = Val(ArchSizeLbl.Caption)
    ValNew = ValNow + HuffSize
    ArchSizeLbl.Caption = ValNew
    TaskLbl.Caption = "Storing Huffman Compressed File"
    DoEvents
    F = FreeFile
    Open TmpFileHuff For Binary As #F 'Open the file, in binary
        ReDim ByteArr(0 To LOF(F) - 1) 'clear the bytearray
        Get #F, , ByteArr() 'Get the file in our memory
    Close #F 'Close the file

    PropBag.WriteProperty "File" & x, ByteArr() 'Put the file in the propertybag, to get it back in the self extract exe.
    PropBag.WriteProperty "File" & x & "Name", "H" & SrcFileNameLbl.Caption 'Put the original filename, so the self extracter knows how to name the filename.
End Sub
Private Sub StoreOriginal(x As Integer, CFileName As String, Origsize As Long)
    'store the original file
    ListView1.ListItems(x).SmallIcon = 4 'show user which storage method selected
    ListView1.ListItems(x).ListSubItems(2).Text = Origsize
    ValNow = Val(ArchSizeLbl.Caption)
    ValNew = ValNow + Origsize
    ArchSizeLbl.Caption = ValNew
    
    F = FreeFile
    Open CFileName For Binary As #F 'Open the file, in binary
        ReDim ByteArr(0 To LOF(F) - 1) 'clear the bytearray
        Get #F, , ByteArr() 'Get the file in our memory
    Close #F 'Close the file

    PropBag.WriteProperty "File" & x, ByteArr() 'Put the file in the propertybag, to get it back in the self extract exe.
    PropBag.WriteProperty "File" & x & "Name", "O" & SrcFileNameLbl.Caption 'Put the original filename, so the self extracter knows how to name the filename.
    'the O prefix tells the extractor to copy the file as is
End Sub

Private Sub Command2_Click()
ListView1.ListItems.Remove ListView1.SelectedItem.Index
End Sub

Private Sub Command3_Click()
    Dim MyFile, Msg
    MyFile = fncGetFileNametoOpen("Add Any File", "All Files|*.*")
    If FileSelected Then
        'FileSelected returned true
            If FileLen(MyFile) = 0 Then
                'stop a subscript out of range if you try to deal with 0 length file
                Msg = "Can't deal with files of zero length!..." & vbCrLf & vbCrLf
                Msg = Msg & "If you need a zero length file... have setup create it!"
                Title = "Zero Length File Encountered!"
                MsgBox Msg, vbOKOnly + vbExclamation, Title
                Exit Sub
            End If
         
         Set itmX = ListView1.ListItems.Add()
         itmX.Text = MyFile
         itmX.SubItems(1) = FileLen(MyFile)
         itmX.SubItems(2) = "???" 'compressed size isn't known yet
         itmX.SubItems(3) = "???" 'stored size isn't known yet
         TtlFileSize = TtlFileSize + FileLen(MyFile)
         TtlFiles = ListView1.ListItems.Count
         frmMain.Caption = App.Title & " (" & TtlFiles & " Files - Size: " & TtlFileSize & ")"
         TtlFileLbl.Caption = TtlFiles
         RemFileLbl.Caption = TtlFiles
        'get the file name without path for extraction use & store for now
        MyFileTitle = Mid(MyFile, InStrRev(MyFile, "\") + 1)
        itmX.SubItems(4) = MyFileTitle
        SizeNow = UpDateSizes
        ArchSizeLbl.Caption = SizeNow
    Else
        'FileSelected returned false - canceled
        'don't do anything unless you want to
        'MSg = "The CANCEL button was selected."
        'MsgBox MSg
    End If
End Sub

Private Sub Command4_Click()
    
    Dim MyFile, Msg
    MyFile = fncGetFileNametoOpen("Select Installer Exe File", "Exe Files|*.exe")
    If FileSelected Then
        'FileSelected returned true
        TemplateFileName = MyFile
        Text1(2).Text = TemplateFileName
        SizeNow = UpDateSizes
        ArchSizeLbl.Caption = SizeNow
    Else
        'FileSelected returned false - canceled
        'MSg = "The CANCEL button was selected."
        'MsgBox MSg
        'don't do anything unless you want to
    End If
     
End Sub
Private Function UpDateSizes()
    'this function calculates the maximum byte size of an uncompressed archive
    UpDateSizes = 0
    If TemplateFileName <> "" Then
        UpDateSizes = FileLen(TemplateFileName)
    End If
    If TtlFiles <> 0 Then
        For x = 1 To TtlFiles
            UpDateSizes = UpDateSizes + Val(ListView1.ListItems(x).ListSubItems(1).Text)
        Next
    End If
    
End Function
Private Sub UpDateLast()
    'this sub calculates the maximum byte size of a compressed/mixed archive
    SizeNow = 0
    SizeNow = FileLen(OutPutFileName)
    ArchSizeLbl.Caption = SizeNow
End Sub
Private Sub Command5_Click()
    On Error GoTo ErrH
    '--> This sub by Written by Allen S. Donker... see credits
    '--- browse for folder
    Dim sPath As String

            '   Open the Browse Folder dialog box and
            '   return the folder selected
    sPath = SelectFolder(Me, "Select folder")
  
  
            '   If no folder was selected, exit here
    If Len(sPath) = 0 Then
        Text1(0).Text = ""
        Exit Sub
    Else
        If Right(sPath, 1) <> "\" Then sPath = sPath & "\"
        Text1(0).Text = sPath
    End If

Exit Sub
    
ErrH:
    MsgBox Err.Number & Chr(10) & Err.Description
End Sub

Private Sub Form_Load()

    Set Huffman = New clsHuffman
    TtlFileSize = 0: TtlFiles = 0
    frmMain.Caption = App.Title & " (" & TtlFiles & " Files - Size: " & TtlFileSize & ")"
    'frmMain.Caption = App.Title & " [" & TtlFiles & " Files - Total File Size: " & TtlFileSize & "]"
    For x = 0 To 1
        Text1(x).Text = ""
    Next
    TemplateFileName = App.path & "\NSSA.EXE"
    FirstTime = True
'    Text1(2).Text = TemplateFileName
    Text1(0).Text = App.path & "\"
    Text1(1).Text = "archive.exe"
End Sub


Private Sub Huffman_Progress(Procent As Integer)

  FileProgressLbl.Caption = Procent & "%"
  prg1.Value = Procent
  DoEvents
  
End Sub

