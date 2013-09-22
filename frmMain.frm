VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{1ABC71B2-B0F7-4C1D-9870-3DED8934B20B}#1.0#0"; "prjXTab.ocx"
Begin VB.Form frmMain 
   Caption         =   "nExtract"
   ClientHeight    =   4005
   ClientLeft      =   225
   ClientTop       =   525
   ClientWidth     =   10395
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4005
   ScaleWidth      =   10395
   StartUpPosition =   3  'Windows Default
   Begin prjXTab.XTab nTab 
      Height          =   3375
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   5953
      TabCaption(0)   =   "Files"
      TabPicture(0)   =   "frmMain.frx":29C12
      TabContCtrlCnt(0)=   7
      Tab(0)ContCtrlCap(1)=   "cmdAdd"
      Tab(0)ContCtrlCap(2)=   "cmdDel"
      Tab(0)ContCtrlCap(3)=   "cmdClear"
      Tab(0)ContCtrlCap(4)=   "cmdSetDestination"
      Tab(0)ContCtrlCap(5)=   "cmdSelectAll"
      Tab(0)ContCtrlCap(6)=   "cmdSelectNone"
      Tab(0)ContCtrlCap(7)=   "lvwFiles"
      TabCaption(1)   =   "Configure"
      TabPicture(1)   =   "frmMain.frx":2A464
      TabContCtrlCnt(1)=   13
      Tab(1)ContCtrlCap(1)=   "txtFileToRunAfterExtract"
      Tab(1)ContCtrlCap(2)=   "chkCompressContents"
      Tab(1)ContCtrlCap(3)=   "chkPropmptDestination"
      Tab(1)ContCtrlCap(4)=   "chkRunFileAfter"
      Tab(1)ContCtrlCap(5)=   "cmdBrowse"
      Tab(1)ContCtrlCap(6)=   "Text11"
      Tab(1)ContCtrlCap(7)=   "txtDefaultDestination"
      Tab(1)ContCtrlCap(8)=   "txtProjectName"
      Tab(1)ContCtrlCap(9)=   "Text10"
      Tab(1)ContCtrlCap(10)=   "Label7"
      Tab(1)ContCtrlCap(11)=   "Label11"
      Tab(1)ContCtrlCap(12)=   "lblProjectName"
      Tab(1)ContCtrlCap(13)=   "Label10"
      TabCaption(2)   =   "Make"
      TabPicture(2)   =   "frmMain.frx":2B1B6
      TabContCtrlCnt(2)=   21
      Tab(2)ContCtrlCap(1)=   "ctlTotalProgress"
      Tab(2)ContCtrlCap(2)=   "cmdRipToIso"
      Tab(2)ContCtrlCap(3)=   "cmdBurnISO"
      Tab(2)ContCtrlCap(4)=   "cmdCreateArchive"
      Tab(2)ContCtrlCap(5)=   "prgMakeProgress"
      Tab(2)ContCtrlCap(6)=   "Label9"
      Tab(2)ContCtrlCap(7)=   "Label3"
      Tab(2)ContCtrlCap(8)=   "SrcFileNameLbl"
      Tab(2)ContCtrlCap(9)=   "Label5"
      Tab(2)ContCtrlCap(10)=   "SrcSizeLbl"
      Tab(2)ContCtrlCap(11)=   "Label6"
      Tab(2)ContCtrlCap(12)=   "ComprSizeLbl"
      Tab(2)ContCtrlCap(13)=   "Label8"
      Tab(2)ContCtrlCap(14)=   "TaskLbl"
      Tab(2)ContCtrlCap(15)=   "TtlFileLbl"
      Tab(2)ContCtrlCap(16)=   "Label12"
      Tab(2)ContCtrlCap(17)=   "RemFileLbl"
      Tab(2)ContCtrlCap(18)=   "Label2"
      Tab(2)ContCtrlCap(19)=   "ArchSizeLbl"
      Tab(2)ContCtrlCap(20)=   "Label4"
      Tab(2)ContCtrlCap(21)=   "FileProgressLbl"
      TabTheme        =   2
      ShowFocusRect   =   0   'False
      InActiveTabBackStartColor=   -2147483626
      InActiveTabBackEndColor=   -2147483626
      InActiveTabForeColor=   -2147483631
      BeginProperty ActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty InActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OuterBorderColor=   -2147483628
      TabStripBackColor=   -2147483626
      DisabledTabBackColor=   -2147483633
      DisabledTabForeColor=   -2147483627
      PictureSize     =   1
      Begin VB.TextBox txtFileToRunAfterExtract 
         Height          =   285
         Left            =   -72960
         TabIndex        =   44
         Top             =   1200
         Width           =   7095
      End
      Begin nExtract.ctlProgressBar ctlTotalProgress 
         Height          =   300
         Left            =   -74880
         TabIndex        =   42
         Top             =   2640
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   529
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
         Color           =   0
      End
      Begin nExtract.ctlXPButton cmdRipToIso 
         Height          =   375
         Left            =   -66360
         TabIndex        =   41
         Top             =   1320
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         BTYPE           =   1
         TX              =   "Rip to ISO"
         ENAB            =   0   'False
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
         MICON           =   "frmMain.frx":2BA08
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin nExtract.ctlXPButton cmdBurnISO 
         Height          =   375
         Left            =   -66360
         TabIndex        =   40
         Top             =   840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         BTYPE           =   1
         TX              =   "Burn ISO"
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
         MICON           =   "frmMain.frx":2BA24
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin nExtract.ctlXPButton cmdCreateArchive 
         Default         =   -1  'True
         Height          =   375
         Left            =   -66360
         TabIndex        =   23
         Top             =   1800
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         BTYPE           =   1
         TX              =   "&Make Archive"
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
         MICON           =   "frmMain.frx":2BA40
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin nExtract.ctlProgressBar prgMakeProgress 
         Height          =   300
         Left            =   -74880
         TabIndex        =   22
         Top             =   2280
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   529
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
         Color           =   0
      End
      Begin VB.CheckBox chkCompressContents 
         Appearance      =   0  'Flat
         Caption         =   "Compression (Slower)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -72960
         TabIndex        =   21
         Top             =   2280
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox chkPropmptDestination 
         Appearance      =   0  'Flat
         Caption         =   "Prompt Destination"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -72960
         TabIndex        =   20
         Top             =   2535
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox chkRunFileAfter 
         Appearance      =   0  'Flat
         Caption         =   "Run file after extract"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -72960
         TabIndex        =   19
         Top             =   2805
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin nExtract.ctlXPButton cmdBrowse 
         Height          =   300
         Left            =   -65760
         TabIndex        =   14
         Top             =   480
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   529
         BTYPE           =   1
         TX              =   "&Browse"
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
         MICON           =   "frmMain.frx":2BA5C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   -72960
         TabIndex        =   13
         Top             =   840
         Width           =   7095
      End
      Begin VB.TextBox txtDefaultDestination 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -72960
         TabIndex        =   12
         Top             =   1920
         Width           =   7095
      End
      Begin VB.TextBox txtProjectName 
         Height          =   285
         Left            =   -72960
         TabIndex        =   11
         Top             =   1560
         Width           =   7095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   -72960
         TabIndex        =   10
         Top             =   480
         Width           =   7095
      End
      Begin nExtract.ctlXPButton cmdAdd 
         Height          =   420
         Left            =   120
         TabIndex        =   9
         Top             =   2280
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   741
         BTYPE           =   1
         TX              =   "&Add"
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
         MICON           =   "frmMain.frx":2BA78
         PICN            =   "frmMain.frx":2BA94
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin nExtract.ctlXPButton cmdDel 
         Height          =   420
         Left            =   1080
         TabIndex        =   8
         Top             =   2280
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   741
         BTYPE           =   1
         TX              =   "&Del"
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
         MICON           =   "frmMain.frx":556B6
         PICN            =   "frmMain.frx":556D2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin nExtract.ctlXPButton cmdClear 
         Height          =   420
         Left            =   1920
         TabIndex        =   7
         Top             =   2280
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   741
         BTYPE           =   1
         TX              =   "&Clear"
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
         MICON           =   "frmMain.frx":7F2F4
         PICN            =   "frmMain.frx":7F310
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin nExtract.ctlXPButton cmdSetDestination 
         Height          =   420
         Left            =   2880
         TabIndex        =   6
         Top             =   2280
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   741
         BTYPE           =   1
         TX              =   "Set Destination"
         ENAB            =   0   'False
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
         MICON           =   "frmMain.frx":A8F32
         PICN            =   "frmMain.frx":A8F4E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin nExtract.ctlXPButton cmdSelectAll 
         Height          =   420
         Left            =   4560
         TabIndex        =   5
         Top             =   2280
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   741
         BTYPE           =   1
         TX              =   "&Select All"
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
         MICON           =   "frmMain.frx":D2B70
         PICN            =   "frmMain.frx":D2B8C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin nExtract.ctlXPButton cmdSelectNone 
         Height          =   420
         Left            =   6240
         TabIndex        =   4
         Top             =   2280
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   741
         BTYPE           =   1
         TX              =   "Select None"
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
         MICON           =   "frmMain.frx":FC7AE
         PICN            =   "frmMain.frx":FC7CA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComctlLib.ListView lvwFiles 
         DragIcon        =   "frmMain.frx":1263EC
         Height          =   1695
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   2990
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         OLEDragMode     =   1
         OLEDropMode     =   1
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList2"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         OLEDragMode     =   1
         OLEDropMode     =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "File"
            Object.Width           =   5292
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
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Destination"
            Object.Width           =   1235
         EndProperty
      End
      Begin VB.Label Label7 
         Caption         =   "File to run after Extract:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   43
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "File Count:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   39
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "File:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   38
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label SrcFileNameLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "None"
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   -73680
         TabIndex        =   37
         Top             =   360
         Width           =   6735
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Size:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   36
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label SrcSizeLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   -73680
         TabIndex        =   35
         Top             =   600
         Width           =   8655
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "New Size:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   34
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label ComprSizeLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   -73680
         TabIndex        =   33
         Top             =   840
         Width           =   8775
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Task:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   32
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label TaskLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "None"
         ForeColor       =   &H00808080&
         Height          =   225
         Left            =   -73680
         TabIndex        =   31
         Top             =   1320
         Width           =   8775
      End
      Begin VB.Label TtlFileLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   -73680
         TabIndex        =   30
         Top             =   1560
         Width           =   8775
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Remaining:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   29
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label RemFileLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   -73680
         TabIndex        =   28
         Top             =   1800
         Width           =   7215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Archive Size:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   27
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label ArchSizeLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   -73680
         TabIndex        =   26
         Top             =   2040
         Width           =   7215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Progress:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   25
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label FileProgressLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "0%"
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   -73680
         TabIndex        =   24
         Top             =   1080
         Width           =   8775
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Output File"
         Height          =   195
         Index           =   1
         Left            =   -74880
         TabIndex        =   18
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label11 
         Caption         =   "Default Destination:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   17
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label lblProjectName 
         Caption         =   "Project Name:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   16
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Output Directory:"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   0
         Left            =   -74880
         TabIndex        =   15
         Top             =   480
         Width           =   1560
      End
   End
   Begin nExtract.ctlXPButton cmdAbout 
      Height          =   375
      Left            =   7200
      TabIndex        =   0
      Top             =   3480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   1
      TX              =   "A&bout"
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
      MICON           =   "frmMain.frx":14FFFE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   120
      Top             =   3360
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
            Picture         =   "frmMain.frx":15001A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":179C3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A385E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1CD480
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1F70A2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin nExtract.ctlXPButton cmdExit 
      Height          =   375
      Left            =   8760
      TabIndex        =   1
      Top             =   3480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   1
      TX              =   "E&xit"
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
      MICON           =   "frmMain.frx":220CC4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   10200
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Menu mnuListViewMenu 
      Caption         =   "<ListView Menu>"
      Visible         =   0   'False
      Begin VB.Menu mnuAddFiles 
         Caption         =   "Add File(s)"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete File(s)"
      End
      Begin VB.Menu mnuClearAllFiles 
         Caption         =   "Clear File(s)"
      End
      Begin VB.Menu mnuSep296326 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSetDestination 
         Caption         =   "Set Destination"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSep320897 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "Select All"
      End
      Begin VB.Menu mnuSelectNone 
         Caption         =   "Select None"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lMessagesVisible As Boolean
Private WithEvents lHuffman As clsHuffman
Attribute lHuffman.VB_VarHelpID = -1
Private lTotalFileSize As Currency, lTotalFiles As Integer, lFileTitle As String, lCFileName As String
Private lOutPutFileName As String, lTempFileHuffman As String, lTempFileLZSS As String, lTempFileName As String
Private lPropBag As New PropertyBag
Private lByteArray() As Byte
Private lFirstRun As Boolean

Public Sub ShowMessage(lMessage As String)
On Local Error Resume Next
If lMessagesVisible = True Then
    frmMessages.AddToListView lMessage
Else
    lMessagesVisible = True
    frmMessages.Show
    frmMessages.AddToListView lMessage
End If
End Sub

Public Sub AddFileToListView(lFile As String, lDestination As String)
On Local Error Resume Next
Dim lListItem As ListItem, lSizeNow, i As Integer
For i = 1 To lvwFiles.ListItems.count
    If LCase(Trim(lvwFiles.ListItems(i).Text)) = LCase(Trim(lFile)) Then
        MsgBox "File '" & lFile & " Exists in List", vbExclamation
        Exit Sub
    End If
Next i
lFile = Replace(lFile, "\\", "\")
If FileLen(lFile) = 0 Or Len(lDestination) = 0 Then
    ShowMessage "The file '" & lFile & "' could not be added"
    Exit Sub
End If
Set lListItem = lvwFiles.ListItems.Add()
lListItem.Text = lFile
lListItem.SubItems(1) = FileLen(lFile)
lListItem.SubItems(2) = "."
lListItem.SubItems(3) = "."
lListItem.SubItems(5) = lDestination
lListItem.Checked = True
lListItem.SmallIcon = 1
lTotalFileSize = lTotalFileSize + FileLen(lFile)
lTotalFiles = lvwFiles.ListItems.count
frmMain.Caption = App.Title & " (" & lTotalFiles & " Files - Size: " & lTotalFileSize & ")"
TtlFileLbl.Caption = lTotalFiles
RemFileLbl.Caption = lTotalFiles
lFileTitle = Mid(lFile, InStrRev(lFile, "\") + 1)
lListItem.SubItems(4) = lFileTitle
lSizeNow = UpDateSizes
ArchSizeLbl.Caption = lSizeNow
End Sub

Private Sub StoreLZSS(X As Integer, lTempFileLZSS As String, LZSize As Long)
On Local Error Resume Next
Dim ValNew, ValNow, F
lvwFiles.ListItems(X).SmallIcon = 2
lvwFiles.ListItems(X).ListSubItems(2).Text = LZSize
ValNow = Val(ArchSizeLbl.Caption)
ValNew = ValNow + LZSize
ArchSizeLbl.Caption = ValNew
TaskLbl.Caption = "Compressing: LZSS Format"
DoEvents
F = FreeFile
Open lTempFileLZSS For Binary As #F
    ReDim lByteArray(0 To LOF(F) - 1)
    Get #F, , lByteArray()
Close #F
lPropBag.WriteProperty "File" & X, lByteArray()
lPropBag.WriteProperty "File" & X & "Name", "L" & SrcFileNameLbl.Caption
End Sub

Private Sub StoreHuffman(X As Integer, lTempFileHuffman As String, HuffSize As Long)
On Local Error Resume Next
Dim ValNow, ValNew, F
lvwFiles.ListItems(X).SmallIcon = 3
lvwFiles.ListItems(X).ListSubItems(2).Text = HuffSize
ValNow = Val(ArchSizeLbl.Caption)
ValNew = ValNow + HuffSize
ArchSizeLbl.Caption = ValNew
TaskLbl.Caption = "Compressing: Huffman Format"
DoEvents
F = FreeFile
Open lTempFileHuffman For Binary As #F
    ReDim lByteArray(0 To LOF(F) - 1)
    Get #F, , lByteArray()
Close #F
lPropBag.WriteProperty "File" & X, lByteArray()
lPropBag.WriteProperty "File" & X & "Name", "H" & SrcFileNameLbl.Caption
End Sub

Private Sub StoreOriginal(X As Integer, lCFileName As String, Origsize As Long)
On Local Error Resume Next
Dim ValNow, ValNew, F
lvwFiles.ListItems(X).SmallIcon = 4
lvwFiles.ListItems(X).ListSubItems(2).Text = Origsize
ValNow = Val(ArchSizeLbl.Caption)
ValNew = ValNow + Origsize
ArchSizeLbl.Caption = ValNew
F = FreeFile
Open lCFileName For Binary As #F
    ReDim lByteArray(0 To LOF(F) - 1)
    Get #F, , lByteArray()
Close #F
lPropBag.WriteProperty "File" & X, lByteArray()
lPropBag.WriteProperty "File" & X & "Name", "O" & SrcFileNameLbl.Caption
End Sub

Private Sub StorePropBagSettings(lPromptRunFile As Boolean, lRunFile As String, lPromptDestination As Boolean, lProjectName As String)
'On Local Error Resume Next
lPropBag.WriteProperty "PromptRunFile", CInt(lPromptRunFile)
lPropBag.WriteProperty "RunFile", lRunFile
lPropBag.WriteProperty "PromptDestination", CInt(lPromptDestination)
lPropBag.WriteProperty "ProjectName", lProjectName
End Sub

Private Sub cmdAbout_Click()
On Local Error Resume Next
frmAbout.Show 1
End Sub

Private Sub cmdAdd_Click()
On Local Error Resume Next
frmAddMultipleFiles.Show
End Sub

Private Sub cmdBrowse_Click()
On Local Error Resume Next
Dim msg As String
msg = SelectDirectory(Me, "Select folder")
If Len(Trim(msg)) <> 0 Then
    If Right(msg, 1) <> "\" Then msg = msg & "\"
    Text1(0).Text = msg
End If
Exit Sub
ErrH:
    MsgBox Err.Number & Chr(10) & Err.Description
End Sub

Private Sub cmdBurnISO_Click()
frmBurnISO.Show 1
End Sub

Private Sub cmdClear_Click()
On Local Error Resume Next
lvwFiles.ListItems.Clear
End Sub

Private Sub cmdCreateArchive_Click()
'On Error GoTo ErrorHandler
Dim msg As String, F, todo, savings, Outputs, mbox As VbMsgBoxResult
If Len(Trim(lTempFileName)) = 0 Then
    MsgBox "No Template EXE file set; Compile aborted.", vbExclamation
    Exit Sub
End If
If Len(Text1(0).Text) = 0 Then
    MsgBox "The output directory is not set; Compile aborted.", vbExclamation
    Exit Sub
End If
If Len(Text1(1).Text) = 0 Then
    MsgBox "The output file is not set; Compile aborted.", vbExclamation
    Exit Sub
End If
If lTotalFiles = 0 Then
    MsgBox "No files have been selected; Compile aborted.", vbOKOnly + vbExclamation
    Exit Sub
End If
If Len(Text1(0).Text) <> 0 And Right(UCase(Text1(0).Text), 1) <> "\" Then Text1(0).Text = Text1(0).Text & "\"
If Len(Text1(1).Text) <> 0 And Right(LCase(Text1(1).Text), 4) <> ".exe" Then Text1(1).Text = Text1(1).Text & ".exe"
If lFirstRun = False Then
    msg = "Critical error. You must restart " & App.Title
    MsgBox msg, vbOKOnly + vbExclamation, App.Title
    Exit Sub
Else
    lFirstRun = False
End If
Text1(0).Enabled = False
Text1(1).Enabled = False
cmdDel.Enabled = False
cmdAdd.Enabled = False
cmdBrowse.Enabled = False
cmdClear.Enabled = False
cmdCreateArchive.Enabled = False
lPropBag.WriteProperty "FileCount", lTotalFiles
lOutPutFileName = Text1(0).Text & Text1(1).Text
lTempFileHuffman = Text1(0).Text & "compareHuff.tmp"
lTempFileLZSS = Text1(0).Text & "compareLZSS.tmp"
F = FreeFile
Open lOutPutFileName For Output As #F
Close #F
FileCopy lTempFileName, lOutPutFileName
Dim Origsize As Long, HuffSize As Long, LZSize As Long, X As Integer, Smallest As Long
Me.MousePointer = 11
todo = Val(TtlFileLbl.Caption)
RemFileLbl.Caption = todo
ctlTotalProgress.Max = lTotalFiles
For X = 1 To lTotalFiles
    ctlTotalProgress.Value = X
    SrcFileNameLbl.Caption = lvwFiles.ListItems(X).ListSubItems(4).Text
    SrcSizeLbl.Caption = lvwFiles.ListItems(X).ListSubItems(1).Text
    ComprSizeLbl.Caption = "0"
    FileProgressLbl.Caption = "0%"
    lvwFiles.ListItems(X).SmallIcon = 1
    lCFileName = lvwFiles.ListItems(X).Text
    Origsize = Val(lvwFiles.ListItems(X).ListSubItems(1).Text)
    Smallest = Origsize
    If lvwFiles.ListItems(X).Checked = False Or chkCompressContents.Value = 0 Then
        TaskLbl.Caption = "Working with file(s)"
        DoEvents
        StoreOriginal X, lCFileName, Origsize
        lvwFiles.ListItems(X).ListSubItems(3).Text = 0
        GoTo SkipIt
    End If
    If lvwFiles.ListItems(X).Checked = False Then GoTo SkipHuffman
    TaskLbl.Caption = "Compressing: " & SrcFileNameLbl.Caption & " (Huffman Format)"
    ComprSizeLbl.Caption = "0"
    FileProgressLbl.Caption = "0%"
    DoEvents
    F = FreeFile
    Open lTempFileHuffman For Output As #F
    Close #F
    HuffSize = 0
    Call lHuffman.EncodeFile(lCFileName, lTempFileHuffman)
    HuffSize = FileLen(lTempFileHuffman)
    ComprSizeLbl.Caption = HuffSize
    If Origsize <= HuffSize Then
        Smallest = Origsize
    Else
        Smallest = HuffSize
    End If
SkipHuffman:
    TaskLbl.Caption = "Compressing: " & SrcFileNameLbl.Caption & " (LZSS Format)"
    DoEvents
    F = FreeFile
    Open lCFileName For Binary As #F
        ReDim OriginalArray(0 To LOF(F) - 1)
        Get #F, , OriginalArray()
    Close #F
    Call CompressLZSS3(OriginalArray)
    F = FreeFile
    Open lTempFileLZSS For Output As #F
    Close #F
    F = FreeFile
    Open lTempFileLZSS For Binary As #F
        Put #F, , OriginalArray()
    Close #F
    LZSize = FileLen(lTempFileLZSS)
    ComprSizeLbl.Caption = LZSize
    If Smallest <= LZSize Then
        Smallest = Origsize
    Else
        Smallest = LZSize
    End If
    TaskLbl.Caption = "Working with file(s)"
    DoEvents
    If Smallest = Origsize Then
        StoreOriginal X, lCFileName, Origsize
        lvwFiles.ListItems(X).ListSubItems(3).Text = 0
        GoTo SkipIt
    End If
    If Smallest = HuffSize Then
        StoreHuffman X, lTempFileHuffman, HuffSize
        GoTo SkipIt
    End If
    If Smallest = LZSize Then
        StoreLZSS X, lTempFileLZSS, LZSize
    End If
    savings = 0
    savings = Origsize - Smallest
    lvwFiles.ListItems(X).ListSubItems(3).Text = savings
SkipIt:
    StorePropBagSettings CBool(chkRunFileAfter.Value), txtFileToRunAfterExtract.Text, CBool(chkPropmptDestination.Value), txtProjectName.Text
    todo = todo - 1
    RemFileLbl.Caption = todo
    DoEvents
Next
TaskLbl.Caption = "Make EXE Complete"
F = FreeFile
Open lOutPutFileName For Binary As #F
    Dim BeginPos As Long
    BeginPos = LOF(F)
    Seek #F, LOF(F)
    Put #F, , lPropBag.Contents
    Put #F, , BeginPos
Close #F
F = FreeFile
Open lTempFileHuffman For Output As #F
Close #F
Kill lTempFileHuffman
F = FreeFile
Open lTempFileLZSS For Output As #F
Close #F
Kill lTempFileLZSS
UpDateLast
Me.MousePointer = 0
savings = 0
For X = 1 To lTotalFiles
    savings = savings + Val(lvwFiles.ListItems(X).ListSubItems(3).Text)
Next X
Outputs = FileLen(lOutPutFileName)
msg = "Make archive operation completed" & vbCrLf & vbCrLf & "New Size: " & Outputs & vbCrLf & "Difference: " & savings & vbCrLf & vbCrLf & "Would you like to test this distribution now? To test click 'Yes', To exit, click 'No'."
mbox = MsgBox(msg, vbInformation + vbYesNo, App.Title)
If mbox = vbYes Then
    Shell lOutPutFileName
    Unload Me
    End
Else
    Unload Me
    End
End If
Exit Sub
ErrorHandler:
    Close
    MsgBox Err.Description, vbExclamation
    End
End Sub

Private Function UpDateSizes()
On Local Error Resume Next
Dim i As Integer
UpDateSizes = 0
If lTempFileName <> "" Then
    UpDateSizes = FileLen(lTempFileName)
End If
If lTotalFiles <> 0 Then
    For i = 1 To lTotalFiles
        UpDateSizes = UpDateSizes + Val(lvwFiles.ListItems(i).ListSubItems(1).Text)
    Next
End If
If Err.Number <> 0 Then
    MsgBox Err.Description
    Err.Clear
End If
End Function

Private Sub UpDateLast()
On Local Error Resume Next
Dim l As Long
l = FileLen(lOutPutFileName)
ArchSizeLbl.Caption = l
End Sub

Private Sub cmdDefaultDestination_Click()

End Sub

Private Sub cmdDel_Click()
On Local Error GoTo ErrHandler
Dim i As Integer, b As Boolean
Do Until b = True
nAgain:
    If lvwFiles.ListItems.count <> 0 Then
        For i = 1 To lvwFiles.ListItems.count
            If lvwFiles.ListItems(i).Selected = True Then
                lvwFiles.ListItems.Remove i
                GoTo nAgain
            End If
            b = True
        Next i
    Else
        Exit Do
    End If
Loop
Exit Sub
ErrHandler:
    MsgBox Err.Description & " " & i
    Err.Clear
End Sub

Private Sub cmdExit_Click()
On Local Error Resume Next
Unload Me
End
End Sub

Private Sub cmdSelectAll_Click()
On Local Error Resume Next
Dim i As Integer
For i = 1 To lvwFiles.ListItems.count
    lvwFiles.ListItems(i).Selected = True
Next i
lvwFiles.SetFocus
End Sub

Private Sub cmdSelectNone_Click()
On Local Error Resume Next
Dim i As Integer
For i = 1 To lvwFiles.ListItems.count
    lvwFiles.ListItems(i).Selected = False
Next i
lvwFiles.SetFocus
End Sub

Private Sub Command1_Click()
frmBurnISO.Show
End Sub

Private Sub Form_Load()
On Local Error Resume Next
Dim i As Integer
Set lHuffman = New clsHuffman
lTotalFileSize = 0: lTotalFiles = 0
frmMain.Caption = App.Title & " (" & lTotalFiles & " Files - Size: " & lTotalFileSize & ")"
For i = 0 To 1
    Text1(i).Text = ""
Next i
lTempFileName = App.Path & "\NSSA.EXE"
lFirstRun = True
Text1(0).Text = App.Path & "\"
End Sub

Private Sub Form_Resize()
On Local Error Resume Next
cmdAbout.Top = Me.ScaleHeight - 475
cmdAbout.Left = Me.ScaleWidth - 3100
cmdExit.Top = Me.ScaleHeight - 475
cmdExit.Left = Me.ScaleWidth - 1570
nTab.Width = Me.ScaleWidth
nTab.Height = Me.ScaleHeight - 600
Line1.X2 = Me.ScaleWidth
Line1.Y1 = Me.ScaleHeight - 550
Line1.Y2 = Me.ScaleHeight - 550
Select Case nTab.ActiveTab
Case 0
    cmdAdd.Top = Me.ScaleHeight - 1000
    cmdAdd.Top = Me.ScaleHeight - 1000
    cmdDel.Top = Me.ScaleHeight - 1000
    cmdClear.Top = Me.ScaleHeight - 1000
    cmdSetDestination.Top = Me.ScaleHeight - 1000
    cmdSelectAll.Top = Me.ScaleHeight - 1000
    cmdSelectNone.Top = Me.ScaleHeight - 1000
    lvwFiles.Height = Me.ScaleHeight - 1500
    lvwFiles.Width = Me.ScaleWidth - 250
Case 1
    txtProjectName.Width = Me.ScaleWidth - 3000
    Text1(0).Width = Me.ScaleWidth - 3000
    Text1(1).Width = Me.ScaleWidth - 3000
    txtDefaultDestination.Width = Me.ScaleWidth - 3000
    txtFileToRunAfterExtract.Width = Me.ScaleWidth - 3000
    cmdBrowse.Left = Me.ScaleWidth - cmdBrowse.Width - 100
'    cmdSetDestination.Left = Me.ScaleWidth - cmdSetDestination.Width - 100
Case 2
    ctlTotalProgress.Top = Me.ScaleHeight - 1000
    ctlTotalProgress.Width = Me.ScaleWidth - 280
    prgMakeProgress.Top = Me.ScaleHeight - 1400
    prgMakeProgress.Width = Me.ScaleWidth - 280
    cmdCreateArchive.Top = Me.ScaleHeight - 1900
    cmdCreateArchive.Left = Me.ScaleWidth - 1600
    cmdBurnISO.Left = Me.ScaleWidth - 1600
    cmdBurnISO.Top = Me.ScaleHeight - 2360
    cmdRipToIso.Left = Me.ScaleWidth - 1600
    cmdRipToIso.Top = Me.ScaleHeight - 2800
End Select
End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub lHuffman_Progress(Procent As Integer)
On Local Error Resume Next
FileProgressLbl.Caption = Procent & "%"
prgMakeProgress.Value = Procent
DoEvents
If Err.Number <> 0 Then
    Err.Clear
End If
End Sub

Private Sub lvwFiles_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error GoTo ErrHandler
If Len(lvwFiles.SelectedItem.Text) <> 0 And Button = 2 Then
    PopupMenu mnuListViewMenu
End If
Exit Sub
ErrHandler:
    If Err.Number <> 0 Then
        Select Case Err.Number
        Case 91
            Err.Clear
        Case Else
            MsgBox Err.Description
        End Select
    End If
End Sub

Private Sub lvwFiles_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error GoTo ErrHandler
Dim i As Integer, c As Integer
c = Data.Files.count
For i = 1 To c
    AddFileToListView Data.Files(i), "\"
Next i
Exit Sub
ErrHandler:
    MsgBox Err.Description, vbExclamation
    Err.Clear
End Sub

Private Sub mnuAddFiles_Click()
On Local Error Resume Next
frmAddMultipleFiles.Show

End Sub

Private Sub mnuClearAllFiles_Click()
On Local Error Resume Next
lvwFiles.ListItems.Clear
End Sub

Private Sub mnuDelete_Click()
'On Local Error GoTo ErrHandler
Dim i As Integer, b As Boolean
Do Until b = True
nAgain:
    If lvwFiles.ListItems.count <> 0 Then
        For i = 1 To lvwFiles.ListItems.count
            If lvwFiles.ListItems(i).Selected = True Then
                lvwFiles.ListItems.Remove i
                GoTo nAgain
            End If
            b = True
        Next i
    Else
        Exit Do
    End If
Loop
Exit Sub
ErrHandler:
    MsgBox Err.Description & " " & i
    Err.Clear
End Sub

Private Sub nTab_Click()
On Local Error Resume Next
Form_Resize
If Err.Number <> 0 Then
    MsgBox Err.Description
    Err.Clear
End If
End Sub
