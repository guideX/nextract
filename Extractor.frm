VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Nexgen Self Extractor"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   5865
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Extractor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   5865
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin NSSA.ctlButton cmdFinish 
      Height          =   315
      Left            =   4800
      TabIndex        =   14
      Top             =   3120
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      Icon            =   "Extractor.frx":29C12
      Style           =   9
      Caption         =   "&Finish"
      iNonThemeStyle  =   0
      Enabled         =   0   'False
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   1
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   0
      RoundedBordersByTheme=   0   'False
   End
   Begin NSSA.ctlButton cmdNext 
      Height          =   315
      Left            =   3720
      TabIndex        =   15
      Top             =   3120
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      Icon            =   "Extractor.frx":29C2E
      Style           =   9
      Caption         =   "&Next >"
      iNonThemeStyle  =   0
      Enabled         =   0   'False
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   1
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   0
      RoundedBordersByTheme=   0   'False
   End
   Begin NSSA.ctlButton cmdBack 
      Height          =   315
      Left            =   2640
      TabIndex        =   16
      Top             =   3120
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      Icon            =   "Extractor.frx":29C4A
      Style           =   9
      Caption         =   "< &Back"
      iNonThemeStyle  =   0
      Enabled         =   0   'False
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   1
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   0
      RoundedBordersByTheme=   0   'False
   End
   Begin NSSA.ctlButton cmdExit 
      Height          =   315
      Left            =   120
      TabIndex        =   17
      Top             =   3120
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      Icon            =   "Extractor.frx":29C66
      Style           =   9
      Caption         =   "&Exit"
      iNonThemeStyle  =   0
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   1
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   0
      RoundedBordersByTheme=   0   'False
   End
   Begin VB.Frame fraSetup 
      BorderStyle     =   0  'None
      Caption         =   "Welcome"
      Height          =   2895
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      Begin VB.Frame fraLicenseAgreement 
         Caption         =   "License Agreement"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   360
         TabIndex        =   19
         Top             =   480
         Width           =   5175
         Begin VB.OptionButton optIAgree 
            Appearance      =   0  'Flat
            Caption         =   "I agree"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   4200
            TabIndex        =   22
            Top             =   1920
            Width           =   855
         End
         Begin VB.OptionButton optDisagree 
            Appearance      =   0  'Flat
            Caption         =   "I disagree"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3120
            TabIndex        =   21
            Top             =   1920
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.TextBox txtLicenseAgreement 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            ForeColor       =   &H00404040&
            Height          =   1575
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   20
            Text            =   "Extractor.frx":29C82
            Top             =   210
            Width           =   4995
         End
         Begin VB.Label lblAgree 
            Caption         =   "Do you agree with these terms?"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   1920
            Width           =   2895
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00404040&
            X1              =   0
            X2              =   5160
            Y1              =   1800
            Y2              =   1800
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00FFFFFF&
            X1              =   0
            X2              =   5160
            Y1              =   1815
            Y2              =   1815
         End
      End
      Begin VB.Image Image3 
         Height          =   240
         Left            =   0
         Top             =   0
         Width           =   240
      End
      Begin VB.Label lblYouRan 
         Caption         =   "Not Initialized"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   5
         Top             =   0
         Width           =   5295
      End
   End
   Begin VB.Frame fraSetup 
      BorderStyle     =   0  'None
      Caption         =   "Finished"
      Height          =   2895
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   5655
      Begin VB.Frame Frame1 
         Caption         =   "Run a file"
         Height          =   2535
         Left            =   360
         TabIndex        =   23
         Top             =   240
         Width           =   5295
         Begin NSSA.ctlButton cmdRunFile 
            Height          =   315
            Left            =   4320
            TabIndex        =   25
            Top             =   2040
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            Icon            =   "Extractor.frx":29C99
            Style           =   9
            Caption         =   "Run File"
            iNonThemeStyle  =   0
            Tooltiptitle    =   ""
            ToolTipIcon     =   0
            ToolTipType     =   1
            ttForeColor     =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaskColor       =   0
            RoundedBordersByTheme=   0   'False
         End
         Begin VB.ListBox lstFiles 
            BackColor       =   &H8000000F&
            Height          =   1695
            IntegralHeight  =   0   'False
            ItemData        =   "Extractor.frx":29CB5
            Left            =   120
            List            =   "Extractor.frx":29CB7
            TabIndex        =   24
            Top             =   240
            Width           =   5025
         End
      End
      Begin VB.Image Image4 
         Height          =   240
         Left            =   0
         Top             =   0
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Self Extract is complete, click 'Finish'"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   0
         Width           =   5175
      End
   End
   Begin VB.Frame fraSetup 
      BorderStyle     =   0  'None
      Caption         =   "Progress"
      Height          =   2895
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   5655
      Begin NSSA.ctlProgressBar XP_ProgressBar1 
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   2520
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   450
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
      End
      Begin VB.Image Image5 
         Height          =   240
         Left            =   0
         Top             =   0
         Width           =   240
      End
      Begin VB.Label TaskLbl 
         Caption         =   "..."
         Height          =   615
         Left            =   360
         TabIndex        =   12
         Top             =   240
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.Label FileProgressLbl 
         Alignment       =   2  'Center
         Caption         =   "0%"
         Height          =   255
         Left            =   4920
         TabIndex        =   11
         Top             =   2520
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Files are being extracted"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   0
         Width           =   5175
      End
   End
   Begin VB.Frame fraSetup 
      BorderStyle     =   0  'None
      Caption         =   "Ready"
      Height          =   2895
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   5655
      Begin VB.Image Image1 
         Height          =   240
         Left            =   0
         Top             =   0
         Width           =   240
      End
      Begin VB.Label lblReady 
         Caption         =   "The destination path has been set and you are ready to extract. Click 'Next'"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         TabIndex        =   8
         Top             =   0
         Width           =   5175
      End
   End
   Begin VB.Frame fraSetup 
      BorderStyle     =   0  'None
      Caption         =   "Location"
      Height          =   2895
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   5655
      Begin NSSA.ctlButton cmdChangeDir 
         Height          =   315
         Left            =   4680
         TabIndex        =   18
         Top             =   600
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         Icon            =   "Extractor.frx":29CB9
         Style           =   9
         Caption         =   "Change"
         iNonThemeStyle  =   0
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   1
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         RoundedBordersByTheme=   0   'False
      End
      Begin VB.Label Label1 
         Caption         =   "To change the location, click 'Change'"
         Height          =   255
         Left            =   360
         TabIndex        =   27
         Top             =   645
         Width           =   3975
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   0
         Top             =   0
         Width           =   240
      End
      Begin VB.Label lblPath 
         Caption         =   "<Path>"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   240
         Width           =   5055
      End
      Begin VB.Label lblInformation 
         Caption         =   "<Info>"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   0
         Width           =   5295
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   5880
      X2              =   0
      Y1              =   3030
      Y2              =   3030
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private lFrameIndex As Integer
Private WithEvents Huffman As clsHuffman
Attribute Huffman.VB_VarHelpID = -1
Private MyExeName As String
Private MyPath As String
Private PropBag As New PropertyBag
Private lAuthor As String
Private lProjectName As String
Private lLicenseAgreement As String
Private lPassword As String
Private lEncryptionPassword As String

Private Sub StartExpanding()
On Error GoTo HndlError
XP_ProgressBar1.Value = 0
FileProgressLbl.Caption = "0%": DoEvents
FileProgressLbl.Visible = True
Me.MousePointer = 11
Dim FileCount As Long
Dim Filename As String
Dim NewFileName As String
Dim ByteArr() As Byte
Dim mbox As VbMsgBoxResult
FileCount = Val(PropBag.ReadProperty("FileCount"))
TaskLbl.Visible = True
TaskLbl.Caption = "Initializing...": DoEvents
For X = 1 To FileCount
BackUp:
    Filename = PropBag.ReadProperty("File" & X & "Name")
    l = Len(Filename)
    NewFileName = MyPath & Mid(Filename, 2, l - 1)
    If DoesFileExist(NewFileName) = True Then
        mbox = MsgBox("File already exists '" & newfile & "'. Would you like to replace this file?", vbYesNo + vbQuestion)
        If mbox = vbNo Then
            MsgBox "File '" & NewFileName & "' was not extracted", vbCritical
            GoTo BackUp
        End If
    End If
    If Right(LCase(Trim(NewFileName)), 4) = ".exe" Then lstFiles.AddItem NewFileName
    Dim ThisFile As String
    ThisFile = MyPath & "NXTRCT0R.TMP"
    If Left(Filename, 1) = "L" Then
        TaskLbl.Caption = "Loading Compressed File...": DoEvents
        F = FreeFile
        Open ThisFile For Output As #F
        Close #F
        F = FreeFile
        Open ThisFile For Binary As #F
            ByteArr() = PropBag.ReadProperty("File" & X)
            Put #F, , ByteArr()
        Close #F
        F = FreeFile
        Open ThisFile For Binary As #F
            ReDim OriginalArray(0 To LOF(F) - 1)
            Get #F, , OriginalArray()
        Close #F
        TaskLbl.Caption = "Decompressing " & Mid(Filename, 2, l - 1) & "...": DoEvents
        Call Decompress_LZSS3(OriginalArray)
        F = FreeFile
        Open NewFileName For Output As #F
        Close #F
        F = FreeFile
        Open NewFileName For Binary As #F
            Put #F, , OriginalArray()
        Close #F
    End If
    If Left(Filename, 1) = "H" Then
        TaskLbl.Caption = "Copying Temporary File...": DoEvents
        F = FreeFile
        Open MyPath & "NXTRCT0R.TMP" For Output As #F
        Close #F
        F = FreeFile
        Open MyPath & "NXTRCT0R.TMP" For Binary As #F
            ByteArr() = PropBag.ReadProperty("File" & X)
            Put #F, , ByteArr()
        Close #F
        TaskLbl.Caption = "Decompressing " & Mid(Filename, 2, l - 1) & "...": DoEvents
        Call Huffman.DecodeFile(ThisFile, NewFileName)
    End If
    If Left(Filename, 1) = "O" Then
        TaskLbl.Caption = "Copying " & Mid(Filename, 2, l - 1) & "...": DoEvents
        F = FreeFile
        Open NewFileName For Output As #F
        Close #F
        F = FreeFile
        Open NewFileName For Binary As #F
            ByteArr() = PropBag.ReadProperty("File" & X)
            Put #F, , ByteArr()
        Close #F
    End If
    pdone = Round((X / FileCount) * 100)
    FileProgressLbl.Caption = pdone & "%": DoEvents
    XP_ProgressBar1.Value = CInt(pdone)
Next
TaskLbl.Caption = "Finished!"
F = FreeFile
Open ThisFile For Output As #F
Close #F
Kill ThisFile
Me.MousePointer = 0
Close
cmdNext_Click
Exit Sub
HndlError:
    Close
    MsgBox "Error: " & Err.Description, vbCritical
    End
End Sub

Private Sub cmdBack_Click()
Dim i As Integer
If cmdBack.Enabled = True Then
    lFrameIndex = lFrameIndex - 1
    For i = 0 To fraSetup.Count - 1
        fraSetup(i).Visible = False
    Next i
    fraSetup(lFrameIndex).Visible = True
End If
End Sub

Private Sub cmdChangeDir_Click()
    On Error GoTo ErrH
    Dim sPath As String
    sPath = SelectFolder(Me, "Select folder")
    If Len(sPath) = 0 Then
        lblPath.Caption = ""
        Exit Sub
    Else
        If Right(sPath, 1) <> "\" Then sPath = sPath & "\"
        lblPath.Caption = sPath
        MyPath = sPath
    End If
Exit Sub
ErrH:
    MsgBox Err.Number & Chr(10) & Err.Description
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdFinish_Click()
If cmdFinish.Enabled = False Then Exit Sub
Unload Me
End Sub

Private Sub cmdNext_Click()
Dim i As Integer
If cmdNext.Enabled = True Then
    If lFrameIndex = fraSetup.Count - 1 Then Exit Sub
    For i = 0 To fraSetup.Count - 1
        fraSetup(i).Visible = False
    Next i
    lFrameIndex = lFrameIndex + 1
    fraSetup(lFrameIndex).Visible = True
    Select Case lFrameIndex
    Case 3
        StartExpanding
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

Private Sub cmdRunFile_Click()
On Local Error Resume Next
Shell lstFiles.Text
End Sub

Public Function DoesFileExist(lFilename As String) As Boolean
'On Local Error Resume Next
Dim Msg As String
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

Private Sub Form_Load()
On Local Error GoTo ErrHandler
Set Huffman = New clsHuffman
Dim BeginPos As Long
Dim tmpByte As Variant
Dim ByteArr() As Byte
lEncryptionPassword = "820huiewfkj4hgru32o0ewyuwe"
Image1.Picture = Me.Icon
Image2.Picture = Me.Icon
Image3.Picture = Me.Icon
Image4.Picture = Me.Icon
Image5.Picture = Me.Icon
MyPath = App.path
If Right(MyPath, 1) <> "\" Then MyPath = MyPath & "\"
MyExeName = MyPath & App.EXEName & ".exe"
lblInformation.Caption = "Ready to extract to:"
lblPath.Caption = MyPath
F = FreeFile
Open MyExeName For Binary As #F
    Get #F, LOF(F) - 3, BeginPos
    Seek #F, BeginPos
    Get #F, , tmpByte
Close #F
ByteArr() = tmpByte
PropBag.Contents = ByteArr()
frmMain.Caption = "Self Extracting Archive"
lProjectName = PropBag.ReadProperty("ProjectName", "")
lAuthor = PropBag.ReadProperty("Author", "")
txtLicenseAgreement.Text = PropBag.ReadProperty("LicenseAgreement", "")
lblReady.Caption = "Enough information has been collected about the target system to extract '" & lProjectName & "'. Click 'Next' to extract file(s)"
lblYouRan.Caption = "You have run the self extracting archive '" & lProjectName & "', click next to continue"
lPassword = PropBag.ReadProperty("Password", "")
If Len(Trim(lPassword)) <> 0 Then
    lPassword = DecodeString(DecodeStr64(lPassword), lEncryptionPassword, True)
    If Len(lPassword) <> 0 Then
        frmPassword.Show 1
    End If
End If
Me.Caption = lProjectName & " by " & lAuthor & " - Self Extracting Archive"
If Len(Trim(txtLicenseAgreement.Text)) = 0 Or Len(Trim(txtLicenseAgreement.Text)) = 1 Or Len(Trim(txtLicenseAgreement.Text)) = 2 Then
    fraLicenseAgreement.Visible = False
    cmdNext.Enabled = True
End If
Me.Visible = True
Exit Sub
ErrHandler:
    MsgBox Err.Description
    Unload Me
End Sub

Private Sub isButton1_Click()
    Close
    End
End Sub

Private Sub isButton4_Click()
    Close
    End
End Sub

Private Sub optDisagree_Click()
cmdNext.Enabled = False
End Sub

Private Sub optIAgree_Click()
cmdNext.Enabled = True
End Sub
