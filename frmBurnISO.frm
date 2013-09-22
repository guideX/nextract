VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmBurnISO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "nExtract - Burn ISO Image WIzard"
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBurnISO.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   172
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   StartUpPosition =   1  'CenterOwner
   Begin nExtract.ctlXPButton cmdExit 
      Height          =   375
      Left            =   120
      TabIndex        =   33
      Top             =   2160
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   1
      TX              =   "&Exit"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmBurnISO.frx":0CCA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin nExtract.ctlXPButton cmdBack 
      Height          =   375
      Left            =   1560
      TabIndex        =   32
      Top             =   2160
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   1
      TX              =   "&Back"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmBurnISO.frx":0CE6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin nExtract.ctlXPButton cmdFinished 
      Height          =   375
      Left            =   3720
      TabIndex        =   31
      Top             =   2160
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   1
      TX              =   "&Finished"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmBurnISO.frx":0D02
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin nExtract.ctlXPButton cmdNext 
      Height          =   375
      Left            =   2640
      TabIndex        =   30
      Top             =   2160
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   1
      TX              =   "&Next"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmBurnISO.frx":0D1E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame fraStep 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Step 5"
      Height          =   2055
      Index           =   4
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   4815
      Begin nExtract.ctlProgressBar XP_ProgressBar1 
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   1560
         Width           =   3495
         _ExtentX        =   6165
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
         Color           =   16750899
      End
      Begin nExtract.ctlXPButton cmdBurnISO 
         Height          =   375
         Left            =   3720
         TabIndex        =   34
         Top             =   1560
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Burn ISO"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   0
         FCOLO           =   16711680
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frmBurnISO.frx":0D3A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblStatus 
         BackStyle       =   0  'Transparent
         Caption         =   "Status: Waiting for Start"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1800
         Width           =   3375
      End
      Begin VB.Label lblStep 
         BackStyle       =   0  'Transparent
         Caption         =   "Step 5 - Start Burning"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   840
         TabIndex        =   17
         Top             =   120
         Width           =   3855
      End
      Begin VB.Image Image5 
         Height          =   720
         Left            =   120
         Picture         =   "frmBurnISO.frx":0D56
         Top             =   120
         Width           =   720
      End
      Begin VB.Label lblStepSM 
         BackStyle       =   0  'Transparent
         Caption         =   "Click 'Burn ISO to start burning"
         Height          =   255
         Index           =   4
         Left            =   840
         TabIndex        =   16
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.Frame fraStep 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Step 2"
      Height          =   2055
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   4815
      Begin nExtract.ctlXPButton cmdSelect 
         Height          =   300
         Left            =   3960
         TabIndex        =   35
         Top             =   1680
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   529
         BTYPE           =   1
         TX              =   "Select"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   0
         FCOLO           =   16711680
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frmBurnISO.frx":298F8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox txtFile 
         Height          =   300
         Left            =   840
         TabIndex        =   3
         Top             =   1680
         Width           =   3045
      End
      Begin VB.Label lblISOFileError 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   840
         TabIndex        =   28
         Top             =   1440
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.Label lblISOFileError2 
         BackStyle       =   0  'Transparent
         Caption         =   "There were problems with the selected file:"
         Height          =   255
         Left            =   840
         TabIndex        =   27
         Top             =   1200
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.Label lblStep 
         BackStyle       =   0  'Transparent
         Caption         =   "Step 2 - Select File"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   20
         Top             =   120
         Width           =   2175
      End
      Begin VB.Image Image2 
         Height          =   720
         Left            =   50
         Picture         =   "frmBurnISO.frx":29914
         Top             =   120
         Width           =   720
      End
      Begin VB.Label lblStepSM 
         BackStyle       =   0  'Transparent
         Caption         =   "Select an ISO file to burn below"
         Height          =   375
         Index           =   2
         Left            =   840
         TabIndex        =   4
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.Frame fraStep 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Step 1"
      Height          =   2055
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      Begin VB.CheckBox chkSkipWelcome 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Skip the welcome message next time"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   840
         TabIndex        =   29
         Top             =   1800
         Width           =   3855
      End
      Begin VB.Label lblStep 
         BackStyle       =   0  'Transparent
         Caption         =   "Step 1 - Welcome"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   21
         Top             =   120
         Width           =   2175
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   90
         Picture         =   "frmBurnISO.frx":524B6
         Top             =   120
         Width           =   720
      End
      Begin VB.Label lblStepSM 
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome to the ISO Image Burner Wizard. Please make sure you have no other programs accessing the CD Burner you wish to burn with."
         Height          =   735
         Index           =   0
         Left            =   840
         TabIndex        =   1
         Top             =   480
         Width           =   3855
      End
   End
   Begin VB.Frame fraStep 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Step 3"
      Height          =   2055
      Index           =   2
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   4815
      Begin MSComctlLib.ListView lvwSettings 
         Height          =   855
         Left            =   840
         TabIndex        =   26
         Top             =   1080
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   1508
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.ComboBox cboDrv 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   720
         Width           =   3855
      End
      Begin VB.Label lblMisc 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "&Misc"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         MouseIcon       =   "frmBurnISO.frx":7B058
         MousePointer    =   99  'Custom
         TabIndex        =   25
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label lblWrite 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "&Write"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         MouseIcon       =   "frmBurnISO.frx":7B1AA
         MousePointer    =   99  'Custom
         TabIndex        =   24
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label lblRead 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "&Read"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         MouseIcon       =   "frmBurnISO.frx":7B2FC
         MousePointer    =   99  'Custom
         TabIndex        =   23
         Top             =   960
         Width           =   615
      End
      Begin VB.Label lblStep 
         BackStyle       =   0  'Transparent
         Caption         =   "Step 3 - Drive Selector"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   840
         TabIndex        =   19
         Top             =   120
         Width           =   3855
      End
      Begin VB.Image Image3 
         Height          =   720
         Left            =   50
         Picture         =   "frmBurnISO.frx":7B44E
         Top             =   120
         Width           =   720
      End
      Begin VB.Label lblStepSM 
         BackStyle       =   0  'Transparent
         Caption         =   "Select the drive below"
         Height          =   375
         Index           =   1
         Left            =   840
         TabIndex        =   6
         Top             =   480
         Width           =   3135
      End
   End
   Begin VB.Frame fraStep 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Step 4"
      Height          =   2055
      Index           =   3
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   4815
      Begin VB.ComboBox cboSpeed 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1680
         Width           =   3855
      End
      Begin VB.CheckBox chkFinalize 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Close Disc"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   840
         TabIndex        =   12
         Top             =   1200
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CheckBox chkEjectDisk 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Eject After Burn"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   840
         TabIndex        =   11
         Top             =   960
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CheckBox chkTestmode 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Test Mode"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   840
         TabIndex        =   10
         Top             =   720
         Width           =   1740
      End
      Begin VB.Label lblStep 
         BackStyle       =   0  'Transparent
         Caption         =   "Step 4 - Start Burning"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   840
         TabIndex        =   18
         Top             =   120
         Width           =   2175
      End
      Begin VB.Image Image4 
         Height          =   720
         Left            =   50
         Picture         =   "frmBurnISO.frx":A3FF0
         Top             =   80
         Width           =   720
      End
      Begin VB.Label lblWriteSpeed 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Speed:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   1720
         Width           =   510
      End
      Begin VB.Label lblStepSM 
         BackStyle       =   0  'Transparent
         Caption         =   "Select optional settings"
         Height          =   255
         Index           =   3
         Left            =   840
         TabIndex        =   9
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   320
      Y1              =   137
      Y2              =   137
   End
End
Attribute VB_Name = "frmBurnISO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private WithEvents cISOCD As FL_CDISOWriter
Attribute cISOCD.VB_VarHelpID = -1
Private lDrvInfo As New FL_DriveInfo
Private lCDInfo As New FL_CDInfo
Private WithEvents lMonitor As FL_DoorMonitor
Attribute lMonitor.VB_VarHelpID = -1
Private lDriveInfo2 As New FL_DriveInfo
Private lCDInfo2 As New FL_CDInfo
Private lTrackInfo As New FL_TrackInfo
Private lSessionInfo As New FL_SessionInfo
Private lReader As New FL_CDReader
Private lPrivateDriveID As String
Private lBuffer() As Byte
Private lDrive As String
Private Type gConfigFiles
    cSettings As String
End Type
Private lConfigFiles As gConfigFiles

Private Sub ShowSpeeds()
On Local Error Resume Next
Dim i As Integer, n() As Integer
n = lDrvInfo.GetWriteSpeeds(strDrvID)
cboSpeed.Clear
For i = LBound(n) To UBound(n)
    cboSpeed.AddItem n(i) & " KB/s (" & (n(i) \ 176) & "x)"
    cboSpeed.ItemData(cboSpeed.ListCount - 1) = n(i)
Next i
cboSpeed.AddItem "Max."
cboSpeed.ItemData(cboSpeed.ListCount - 1) = &HFFFF&
cboSpeed.ListIndex = cboSpeed.ListCount - 1
End Sub

Private Sub cboDrv_Click()
On Local Error Resume Next
RefreshDrive
End Sub

Private Sub cISOCD_ClosingSession()
On Local Error Resume Next
lblStatus.Caption = "Status: Closing Session"
End Sub

Private Sub cISOCD_Finished()
On Local Error Resume Next
lblStatus.Caption = "Status: Finished"
End Sub

Private Sub cISOCD_Progress(Percent As Integer)
On Local Error Resume Next
XP_ProgressBar1.Value = Percent
End Sub

Private Sub cISOCD_StartWriting()
On Local Error Resume Next
lblStatus.Caption = "Status: Writing Track"
End Sub

Private Sub cmdBack_Click()
On Local Error Resume Next
Dim i As Integer, d As Integer
cmdNext.Enabled = True
For i = 0 To fraStep.count
    If fraStep(i).Visible = True Then
        For d = 0 To fraStep.count - 1
            fraStep(d).Visible = False
        Next d
        If i - 1 <> -1 Then
            fraStep(i - 1).Visible = True
        Else
            fraStep(0).Visible = True
        End If
        Exit For
    End If
Next i
End Sub

Private Sub cmdBurnISO_Click()
On Local Error Resume Next
Dim msg As String
If FileLen(txtFile.Text) = 0 Then
    MsgBox "This file has no data, it has a filesize of 0"
    Exit Sub
End If
If txtFile = vbNullString Then
    MsgBox "No ISO image selected.", vbExclamation
    Exit Sub
End If
lCDInfo.GetInfo strDrvID
If FileLen(txtFile) > lCDInfo.Capacity Then
    If MsgBox("Image size exceeds disk capacity." & vbCrLf & "Continue?", vbYesNo Or vbQuestion) = vbNo Then Exit Sub
End If
cISOCD.ISOFile = txtFile
cISOCD.EjectAfterWrite = chkEjectDisk
cISOCD.NextSessionAllowed = Not CBool(chkFinalize)
cISOCD.TestMode = chkTestmode
Me.Hide
'frmDataCDPrg.Show
Select Case cISOCD.WriteISOtoCD(strDrvID)
    Case BURNRET_CLOSE_SESSION: msg = "Could not close session."
    Case BURNRET_CLOSE_TRACK: msg = "Could not cloe track."
    Case BURNRET_FILE_ACCESS: msg = "Failed to access a file."
    Case BURNRET_INVALID_MEDIA: msg = "Invalid medium in drive."
    Case BURNRET_ISOCREATION: msg = "ISO creation failed."
    Case BURNRET_NO_NEXT_WRITABLE_LBA: msg = "Could not get next writable LBA."
    Case BURNRET_NOT_EMPTY: msg = "Disk is finalized."
    Case BURNRET_OK: msg = "Finished."
    Case BURNRET_SYNC_CACHE: msg = "Could not synchronize cache."
    Case BURNRET_WPMP: msg = "Write Parameters Page invalid"
    Case BURNRET_WRITE: msg = "Write error (Buffer Underrun?)"
End Select
MsgBox msg, vbInformation
Me.Show
'Unload frmDataCDPrg
End Sub

Private Sub cmdCancel_Click()

End Sub

Private Sub cmdExit_Click()
On Local Error Resume Next
Unload Me
End Sub

Private Sub cmdFinished_Click()
On Local Error Resume Next
If cmdFinished.Enabled = True Then
    Unload Me
End If
End Sub

Private Sub cmdNext_Click()
On Local Error Resume Next
Dim i As Integer, d As Integer, c As Integer
c = Int(fraStep.count)
For i = 0 To c
    If i <> c And (i + 1) <> c Then
        If fraStep(i).Visible = True Then
            Select Case i
            Case 1
                If Len(txtFile.Text) = 0 Or DoesFileExist(txtFile.Text) = False Then
                    MsgBox "You must select a ISO file to burn", vbInformation
                    txtFile.SetFocus
                    lblISOFileError.Visible = True
                    lblISOFileError2.Visible = True
                    lblISOFileError.Caption = "You must select an ISO file to burn to continue"
                    Exit Sub
                End If
                If FileLen(txtFile.Text) = 0 Then
                    MsgBox "Can not burn '" & GetFileTitle(txtFile.Text) & "'. File is empty", vbExclamation
                    lblISOFileError.Visible = True
                    lblISOFileError2.Visible = True
                    lblISOFileError.Caption = "File Empty"
                    Exit Sub
                End If
            End Select
            lblISOFileError.Visible = False
            lblISOFileError2.Visible = False
            For d = 0 To c - 1
                fraStep(d).Visible = False
            Next d
            fraStep(i + 1).Visible = True
            Exit For
        End If
        cmdNext.Enabled = True
        'cmdFinished.Enabled = False
    Else
        cmdNext.Enabled = False
        'cmdFinished.Enabled = True
        Exit For
    End If
Next i
DoEvents
End Sub

Private Sub cmdSelect_Click()
On Local Error Resume Next
Dim msg As String
If Len(msg) <> 0 Then
    msg = GetFileTitle(txtFile.Text)
    txtFile.Text = Left(txtFile.Text, Len(txtFile.Text) - Len(msg))
    txtFile.Text = OpenDialog(Me, "ISO Files (*.iso)|*.iso|", "Select ISO File", txtFile.Text)
Else
    txtFile.Text = OpenDialog(Me, "ISO Files (*.iso)|*.iso|", "Select ISO File", CurDir)
End If
End Sub

Private Sub RefreshDrive()
On Local Error Resume Next
lDrive = cboDrv.List(cboDrv.ListIndex)
strDrvID = cManager.DrvChr2DrvID(Left$(lDrive, 1))
lPrivateDriveID = cManager.DrvChr2DrvID(Left(lDrive, 1))
DoEvents
ShowSpeeds
DriveInfo
End Sub

Private Sub Form_Load()
On Local Error Resume Next
Dim msg As String
lConfigFiles.cSettings = App.Path & "\settings.ini"
lvwSettings.ColumnHeaders.Add , , "Setting", ReadINI(lConfigFiles.cSettings, "ISOBurnerWizard", "Col1Width", 1000)
lvwSettings.ColumnHeaders.Add , , "Value", ReadINI(lConfigFiles.cSettings, "ISOBurnerWizard", "Col2Width", 1500)
Set cISOCD = New FL_CDISOWriter
ShowDrives
Me.Left = ReadINI(lConfigFiles.cSettings, "ISOBurnerWizard", "Left", 0)
Me.Top = ReadINI(lConfigFiles.cSettings, "ISOBurnerWizard", "Top", 0)
txtFile.Text = ReadINI(lConfigFiles.cSettings, "ISOBurnerWizard", "LastTextFile", "")
msg = ReadINI(lConfigFiles.cSettings, "ISOBurnerWizard", "SkipWelcome", "")
If Len(msg) <> 0 Then
    chkSkipWelcome.Value = Int(Trim(msg))
End If
If chkSkipWelcome.Value = 1 Then
    fraStep(0).Visible = False
    fraStep(1).Visible = True
End If
End Sub

Private Sub ShowDrives()
On Local Error Resume Next
Dim msg() As String, l As Long
msg = GetDriveList(OPT_CDWRITERS)
For l = LBound(msg) To UBound(msg) - 1
    lDrvInfo.GetInfo cManager.DrvChr2DrvID(msg(l))
    With lDrvInfo
        cboDrv.AddItem msg(l) & ": " & .Vendor & " " & .Product & " " & .Revision & " [" & .HostAdapter & ":" & .Target & "]"
    End With
Next l
cboDrv.ListIndex = 0
End Sub

Private Sub AddItem(Text1 As String, text2 As String)
On Local Error Resume Next
Dim i As Integer
lvwSettings.ListItems.Add , , Text1
For i = 1 To lvwSettings.ListItems.count
    If lvwSettings.ListItems(i).Text = Text1 Then
        lvwSettings.ListItems(i).SubItems(1) = text2
    End If
Next i
End Sub

Private Sub DriveInfo()
On Local Error Resume Next
If Not lDrvInfo.GetInfo(lPrivateDriveID) Then
    MsgBox "Could not get drive info.", vbExclamation
    Exit Sub
End If
lvwSettings.ListItems.Clear
With lDrvInfo
    If lblRead.Font.Underline = True Then
        AddItem "Max. read speed", .ReadSpeedMax & " KB/s"
        AddItem "Cur. read speed", .ReadSpeedCur & " KB/s"
        AddItem "CD-R", CBool(.ReadCapabilities And RC_CDR)
        AddItem "CD-RW", CBool(.ReadCapabilities And RC_CDRW)
        AddItem "DVD-ROM", CBool(.ReadCapabilities And RC_DVDROM)
        AddItem "DVD-RAM", CBool(.ReadCapabilities And RC_DVDRAM)
        AddItem "DVD-R", CBool(.ReadCapabilities And RC_DVDR)
        AddItem "DVD+R", CBool(.ReadCapabilities And RC_DVDPR)
        AddItem "DVD+R DL", CBool(.ReadCapabilities And RC_DVDPRDL)
        AddItem "DVD-RW", CBool(.ReadCapabilities And RC_DVDR)
        AddItem "DVD+RW", CBool(.ReadCapabilities And RC_DVDPRW)
        AddItem "C2 Errors", CBool(.ReadCapabilities And RC_C2)
        AddItem "Bar Code", CBool(.ReadCapabilities And RC_BARCODE)
        AddItem "CDDA raw", CBool(.ReadCapabilities And RC_CDDARAW)
        AddItem "CD-Text", CBool(.ReadCapabilities And RC_CDTEXT)
        AddItem "ISRC", CBool(.ReadCapabilities And RC_ISRC)
        AddItem "Mode 2 Form 1", CBool(.ReadCapabilities And RC_MODE2FORM1)
        AddItem "Mode 2 Form 2", CBool(.ReadCapabilities And RC_MODE2FORM2)
        AddItem "Mount Rainer", CBool(.ReadCapabilities And RC_MRW)
        AddItem "Multisession", CBool(.ReadCapabilities And RC_MULTISESSION)
        AddItem "Sub-Channels", CBool(.ReadCapabilities And RC_SUBCHANNELS)
        AddItem "Sub-Channels corrected", CBool(.ReadCapabilities And RC_SUBCHANNELS_CORRECTED)
        AddItem "Sub-Channels from Lead-In", CBool(.ReadCapabilities And RC_SUBCHANNELS_FROM_LEADIN)
    End If
    If lblWrite.Font.Underline = True Then
        AddItem "Max. write speed", .WriteSpeedMax & " KB/s"
        AddItem "Cur. write speed", .WriteSpeedCur & " KB/s"
        AddItem "CD-R", CBool(.WriteCapabilities And WC_CDR)
        AddItem "CD-RW", CBool(.WriteCapabilities And WC_CDRW)
        AddItem "DVD-R", CBool(.WriteCapabilities And WC_DVDR)
        AddItem "DVD+R", CBool(.WriteCapabilities And WC_DVDPR)
        AddItem "DVD+R DL", CBool(.WriteCapabilities And WC_DVDPRDL)
        AddItem "DVD-RW", CBool(.WriteCapabilities And WC_DVDRRW)
        AddItem "DVD+RW", CBool(.WriteCapabilities And WC_DVDPRW)
        AddItem "DVD-RAM", CBool(.WriteCapabilities And WC_DVDRAM)
        AddItem "Mount Rainer", CBool(.WriteCapabilities And WC_MRW)
        AddItem "BURN-Proof", CBool(.WriteCapabilities And WC_BURNPROOF)
        AddItem "Test-Mode", CBool(.WriteCapabilities And WC_TESTMODE)
        AddItem "TAO", CBool(.WriteCapabilities And WC_TAO)
        AddItem "TAO+Test", CBool(.WriteCapabilities And WC_TAO_TEST)
        AddItem "SAO", CBool(.WriteCapabilities And WC_SAO)
        AddItem "SAO+Test", CBool(.WriteCapabilities And WC_SAO_TEST)
        AddItem "DAO/16", CBool(.WriteCapabilities And WC_RAW_16)
        AddItem "DAO/16+Test", CBool(.WriteCapabilities And WC_RAW_16_TEST)
        AddItem "DAO/96", CBool(.WriteCapabilities And WC_RAW_96)
        AddItem "DAO/96+Test", CBool(.WriteCapabilities And WC_RAW_96_TEST)
    End If
    If lblMisc.Font.Underline = True Then
        AddItem "Analog audio playback", .AnalogAudioPlayback
        AddItem "Buffer size", .BufferSizeKB & " KB"
        AddItem "Anti Jitter", .JitterEffectCorrection
        AddItem "Loading mechanism", LoadingMech2Str(.LoadingMechanism)
        AddItem "Lockable", .Lockable
        AddItem "Physical interface", IPh2Str(.PhysicalInterface)
        AddItem "Disk present", .DiscPresent
        AddItem "Drive closed", .DriveClosed
        AddItem "Drive locked", .DriveLocked
        AddItem "Idle timer", .IdleTimer100MS & " ms"
        AddItem "Spindown timer", .SpinDownTimerMS & " ms"
        AddItem "Standby timer", .StandbyTimer100MS & " ms"
    End If
End With
End Sub

Private Function IPh2Str(i As FL_PhysicalInterfaces) As String
On Local Error Resume Next
Select Case i
    Case IF_ATAPI: IPh2Str = "ATAPI"
    Case IF_IEEE: IPh2Str = "IEEE"
    Case IF_SCSI: IPh2Str = "SCSI"
    Case IF_UNKNWN: IPh2Str = "unknown"
    Case IF_USB: IPh2Str = "USB"
End Select
End Function

Private Function LoadingMech2Str(mech As FL_LoadingMech) As String
On Local Error Resume Next
Select Case mech
    Case LOAD_CADDY: LoadingMech2Str = "Caddy"
    Case LOAD_CHANGER: LoadingMech2Str = "Changer"
    Case LOAD_POPUP: LoadingMech2Str = "Popup"
    Case LOAD_TRAY: LoadingMech2Str = "Tray"
    Case LOAD_UNKNWN: LoadingMech2Str = "Unknown"
End Select
End Function

Private Sub Form_Unload(Cancel As Integer)
On Local Error Resume Next
WriteINI lConfigFiles.cSettings, "ISOBurnerWizard", "Left", Me.Left
WriteINI lConfigFiles.cSettings, "ISOBurnerWizard", "Top", Me.Top
WriteINI lConfigFiles.cSettings, "ISOBurnerWizard", "Col1Width", lvwSettings.ColumnHeaders(1).Width
WriteINI lConfigFiles.cSettings, "ISOBurnerWizard", "Col2Width", lvwSettings.ColumnHeaders(2).Width
WriteINI lConfigFiles.cSettings, "ISOBurnerWizard", "LastTextFile", txtFile.Text
WriteINI lConfigFiles.cSettings, "ISOBurnerWizard", "SkipWelcome", Trim(Str(chkSkipWelcome.Value))
End Sub

Private Sub fraStep_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
lblRead.ForeColor = vbBlue
lblWrite.ForeColor = vbBlue
lblMisc.ForeColor = vbBlue
End Sub

Private Sub lblMisc_Click()
On Local Error Resume Next
lblRead.Font.Underline = False
lblWrite.Font.Underline = False
lblMisc.Font.Underline = True
RefreshDrive
End Sub

Private Sub lblMisc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
lblRead.ForeColor = vbBlue
lblWrite.ForeColor = vbBlue
lblMisc.ForeColor = vbGreen
End Sub

Private Sub lblRead_Click()
On Local Error Resume Next
lblRead.Font.Underline = True
lblWrite.Font.Underline = False
lblMisc.Font.Underline = False
RefreshDrive
End Sub

Private Sub lblRead_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
lblRead.ForeColor = vbGreen
lblWrite.ForeColor = vbBlue
lblMisc.ForeColor = vbBlue
End Sub

Private Sub lblWrite_Click()
On Local Error Resume Next
lblRead.Font.Underline = False
lblWrite.Font.Underline = True
lblMisc.Font.Underline = False
RefreshDrive
End Sub

Private Sub lblWrite_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
lblRead.ForeColor = vbBlue
lblWrite.ForeColor = vbGreen
lblMisc.ForeColor = vbBlue
End Sub

Private Sub lvwSettings_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
lblRead.ForeColor = vbBlue
lblWrite.ForeColor = vbBlue
lblMisc.ForeColor = vbBlue
End Sub
