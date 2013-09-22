VERSION 5.00
Object = "{1ABC71B2-B0F7-4C1D-9870-3DED8934B20B}#2.0#0"; "prjXTab.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "nExtract"
   ClientHeight    =   4410
   ClientLeft      =   225
   ClientTop       =   525
   ClientWidth     =   7845
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4410
   ScaleWidth      =   7845
   StartUpPosition =   3  'Windows Default
   Begin nExtract.ctlXPButton cmdAbout 
      Height          =   375
      Left            =   4680
      TabIndex        =   7
      Top             =   3960
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
      MICON           =   "frmMain1.frx":29C12
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin nExtract.ctlXPButton cmdExit 
      Height          =   375
      Left            =   6240
      TabIndex        =   8
      Top             =   3960
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
      MICON           =   "frmMain1.frx":29C2E
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
      Left            =   720
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain1.frx":29C4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain1.frx":5386C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain1.frx":7D48E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain1.frx":A70B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain1.frx":D0CD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain1.frx":FA8F4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   45
      Top             =   0
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   635
      ButtonWidth     =   1376
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Open"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Save"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Clear"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Pass"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Make"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   3840
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
            Picture         =   "frmMain1.frx":124516
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain1.frx":14E138
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain1.frx":177D5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain1.frx":1A197C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain1.frx":1CB59E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin prjXTab.XTab nTab 
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   6165
      TabCaption(0)   =   "Files"
      TabContCtrlCnt(0)=   7
      Tab(0)ContCtrlCap(1)=   "lvwFiles"
      Tab(0)ContCtrlCap(2)=   "cmdAdd"
      Tab(0)ContCtrlCap(3)=   "cmdDel"
      Tab(0)ContCtrlCap(4)=   "cmdClear"
      Tab(0)ContCtrlCap(5)=   "cmdSetDestination"
      Tab(0)ContCtrlCap(6)=   "cmdSelectAll"
      Tab(0)ContCtrlCap(7)=   "cmdSelectNone"
      TabCaption(1)   =   "Configure"
      TabContCtrlCnt(1)=   15
      Tab(1)ContCtrlCap(1)=   "txtPassword"
      Tab(1)ContCtrlCap(2)=   "txtAuthor"
      Tab(1)ContCtrlCap(3)=   "chkCompressContents"
      Tab(1)ContCtrlCap(4)=   "cmdBrowse"
      Tab(1)ContCtrlCap(5)=   "txtOutputFile"
      Tab(1)ContCtrlCap(6)=   "txtLicenseAgreement"
      Tab(1)ContCtrlCap(7)=   "txtProjectName"
      Tab(1)ContCtrlCap(8)=   "txtOutputDirectory"
      Tab(1)ContCtrlCap(9)=   "Label1"
      Tab(1)ContCtrlCap(10)=   "Shape1"
      Tab(1)ContCtrlCap(11)=   "lblAuthor"
      Tab(1)ContCtrlCap(12)=   "lblOutputFile"
      Tab(1)ContCtrlCap(13)=   "lblLicenseAgreement"
      Tab(1)ContCtrlCap(14)=   "lblProjectName"
      Tab(1)ContCtrlCap(15)=   "lblOutputDirectory"
      TabCaption(2)   =   "Make"
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
      Begin MSComctlLib.ListView lvwFiles 
         Height          =   2295
         Left            =   120
         TabIndex        =   44
         Top             =   480
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   4048
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList2"
         SmallIcons      =   "ImageList2"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "File"
            Object.Width           =   4410
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
            Text            =   "FileName"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Destination"
            Object.Width           =   882
         EndProperty
      End
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         IMEMode         =   3  'DISABLE
         Left            =   -72960
         PasswordChar    =   "*"
         TabIndex        =   18
         Top             =   1440
         Width           =   5655
      End
      Begin VB.TextBox txtAuthor 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   -72960
         TabIndex        =   16
         Top             =   1200
         Width           =   5655
      End
      Begin nExtract.ctlProgressBar ctlTotalProgress 
         Height          =   300
         Left            =   -74880
         TabIndex        =   39
         Top             =   2760
         Width           =   7575
         _ExtentX        =   13361
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
         Color           =   32896
      End
      Begin nExtract.ctlXPButton cmdRipToIso 
         Height          =   375
         Left            =   -68760
         TabIndex        =   41
         Top             =   1320
         Visible         =   0   'False
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
         MICON           =   "frmMain1.frx":1F51C0
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
         Left            =   -68760
         TabIndex        =   40
         Top             =   840
         Visible         =   0   'False
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
         MICON           =   "frmMain1.frx":1F51DC
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
         Height          =   375
         Left            =   -68760
         TabIndex        =   42
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
         MICON           =   "frmMain1.frx":1F51F8
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
         TabIndex        =   38
         Top             =   2400
         Width           =   7575
         _ExtentX        =   13361
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
         Color           =   49344
      End
      Begin VB.CheckBox chkCompressContents 
         Appearance      =   0  'Flat
         Caption         =   "Compression (Slower)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -72960
         TabIndex        =   21
         Top             =   2760
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin nExtract.ctlXPButton cmdBrowse 
         Height          =   300
         Left            =   -67200
         TabIndex        =   43
         Top             =   960
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
         MICON           =   "frmMain1.frx":1F5214
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox txtOutputFile 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   -72960
         TabIndex        =   12
         Top             =   720
         Width           =   5655
      End
      Begin VB.TextBox txtLicenseAgreement 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   945
         Left            =   -72960
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Text            =   "frmMain1.frx":1F5230
         Top             =   1680
         Width           =   5655
      End
      Begin VB.TextBox txtProjectName 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   -72960
         TabIndex        =   10
         Top             =   480
         Width           =   5655
      End
      Begin VB.TextBox txtOutputDirectory 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   -72960
         TabIndex        =   14
         Top             =   960
         Width           =   5655
      End
      Begin nExtract.ctlXPButton cmdAdd 
         Height          =   420
         Left            =   120
         TabIndex        =   1
         Top             =   2880
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
         MICON           =   "frmMain1.frx":1F523B
         PICN            =   "frmMain1.frx":1F5257
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
         TabIndex        =   2
         Top             =   2880
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
         MICON           =   "frmMain1.frx":21EE79
         PICN            =   "frmMain1.frx":21EE95
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
         TabIndex        =   3
         Top             =   2880
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
         MICON           =   "frmMain1.frx":248AB7
         PICN            =   "frmMain1.frx":248AD3
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
         Left            =   6120
         TabIndex        =   6
         Top             =   2880
         Visible         =   0   'False
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
         MICON           =   "frmMain1.frx":2726F5
         PICN            =   "frmMain1.frx":272711
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
         Left            =   2880
         TabIndex        =   4
         Top             =   2880
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
         MICON           =   "frmMain1.frx":29C333
         PICN            =   "frmMain1.frx":29C34F
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
         Left            =   4560
         TabIndex        =   5
         Top             =   2880
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
         MICON           =   "frmMain1.frx":2C5F71
         PICN            =   "frmMain1.frx":2C5F8D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label1 
         Caption         =   "Password:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   17
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   2175
         Left            =   -72975
         Top             =   465
         Width           =   5700
      End
      Begin VB.Label lblAuthor 
         Caption         =   "Author:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   15
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "File Count:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   32
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "File:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   22
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label SrcFileNameLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "None"
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   -73680
         TabIndex        =   23
         Top             =   360
         Width           =   6735
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Size:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   24
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label SrcSizeLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   -73680
         TabIndex        =   25
         Top             =   600
         Width           =   8655
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "New Size:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   26
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label ComprSizeLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   -73680
         TabIndex        =   27
         Top             =   840
         Width           =   8775
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Task:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   30
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
         TabIndex        =   33
         Top             =   1560
         Width           =   8775
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Remaining:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   34
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label RemFileLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   -73680
         TabIndex        =   35
         Top             =   1800
         Width           =   7215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Archive Size:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   36
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label ArchSizeLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   -73680
         TabIndex        =   37
         Top             =   2040
         Width           =   7215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Progress:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   28
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label FileProgressLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "0%"
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   -73680
         TabIndex        =   29
         Top             =   1080
         Width           =   8775
      End
      Begin VB.Label lblOutputFile 
         BackStyle       =   0  'Transparent
         Caption         =   "Output File:"
         Height          =   195
         Left            =   -74880
         TabIndex        =   11
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblLicenseAgreement 
         Caption         =   "License Agreement:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   19
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label lblProjectName 
         Caption         =   "Project Name:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   9
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label lblOutputDirectory 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Output Directory:"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   -74880
         TabIndex        =   13
         Top             =   960
         Width           =   1560
      End
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
Private lEncryptionPassword As String
Private lProjectFileName As String

Public Function GetFileTitle(lFileName As String) As String
On Local Error Resume Next
Dim msg() As String
If Len(lFileName) <> 0 Then
    msg = Split(lFileName, "\", -1, vbTextCompare)
    GetFileTitle = msg(UBound(msg))
End If
End Function

Public Sub ShowMessage(lMessage As String, lErrNumber As Integer)
On Local Error Resume Next
If lMessagesVisible = True Then
    frmMessages.AddToListView lMessage, lErrNumber
    frmMessages.Visible = True
Else
    lMessagesVisible = True
    frmMessages.Show
    frmMessages.AddToListView lMessage, Err.Number
End If
End Sub

Public Sub AddFileToListView(lFile As String, lDestination As String)
On Local Error Resume Next
Dim lListItem As ListItem, lSizeNow, i As Integer
For i = 1 To lvwFiles.ListItems.count
    If LCase(Trim(lvwFiles.ListItems(i).Text)) = LCase(Trim(lFile)) Then
        ShowMessage "The file '" & lFile & "' already exists", Err.Number
        Exit Sub
    End If
Next i
lFile = Replace(lFile, "\\", "\")
If FileLen(lFile) = 0 Or Len(lDestination) = 0 Then
    ShowMessage "The file '" & lFile & "' could not be added", Err.Number
    Exit Sub
End If
Set lListItem = lvwFiles.ListItems.Add()
lListItem.Text = lFile
lListItem.SubItems(1) = FileLen(lFile)
lListItem.SubItems(2) = "."
lListItem.SubItems(3) = "."
lListItem.SubItems(5) = lDestination
Select Case LCase(Right(lFile, 4))
Case ".exe"
    lListItem.Checked = False
Case ".mp3"
    lListItem.Checked = False
Case ".wav"
    lListItem.Checked = False
Case ".wmv"
    lListItem.Checked = False
Case ".mpg"
    lListItem.Checked = False
Case "mpeg"
    lListItem.Checked = False
Case ".com"
    lListItem.Checked = False
Case ".dll"
    lListItem.Checked = False
Case ".mdb"
    lListItem.Checked = False
Case ".wma"
    lListItem.Checked = False
Case Else
    lListItem.Checked = True
End Select
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

Private Sub StorePropBagSettings()
On Local Error Resume Next
lPropBag.WriteProperty "ProjectName", txtProjectName.Text
lPropBag.WriteProperty "LicenseAgreement", txtLicenseAgreement.Text
lPropBag.WriteProperty "Author", txtAuthor.Text
lPropBag.WriteProperty "Password", EncodeStr64(EncodeString(txtPassword.Text, lEncryptionPassword, True), 68)
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
    txtOutputDirectory.Text = msg
End If
Exit Sub
ErrH:
    ShowMessage Err.Description, Err.Number
End Sub

Private Sub cmdClear_Click()
On Local Error Resume Next
lvwFiles.ListItems.Clear
End Sub

Public Function DoesFileExist(lFileName As String) As Boolean
On Local Error Resume Next
Dim msg As String
msg = Dir(lFileName)
If msg <> "" Then
    DoesFileExist = True
Else
    DoesFileExist = False
End If
End Function

Public Function DoesFolderExist(lFolder As String) As Boolean
On Local Error GoTo ErrHandler
Dim i As Integer
i = GetAttr(lFolder)
Select Case i
Case 17
    DoesFolderExist = True
Case 0
    DoesFolderExist = True
Case 64
    DoesFolderExist = True
Case 32
    DoesFolderExist = True
Case 16
    DoesFolderExist = True
Case 8
    DoesFolderExist = True
Case 4
    DoesFolderExist = True
Case 2
    DoesFolderExist = True
Case 1
    DoesFolderExist = True
End Select
Exit Function
ErrHandler:
    Err.Clear
End Function

Private Sub cmdCreateArchive_Click()
On Error GoTo ErrorHandler
Dim msg As String, lCurrentAction As String, F, todo, savings, Outputs, mbox As VbMsgBoxResult
If Len(Trim(lTempFileName)) = 0 Then
    ShowMessage "No Template EXE file set: Compile aborted.", Err.Number
    Exit Sub
Else
    If DoesFileExist(Trim(lTempFileName)) = False Then
        mbox = MsgBox("Temp file '" & lTempFileName & "'. Would you like to delete?", vbExclamation + vbYesNo)
        If mbox = vbYes Then Kill lTempFileName: DoEvents
        If DoesFileExist(lTempFileName) = True Then
            MsgBox "Could not delete existing file", vbCritical
            Exit Sub
        End If
    End If
End If
If lTotalFiles = 0 Then
    mbox = MsgBox("No files have been selected. Would you like to select some now?", vbYesNo + vbQuestion)
    If mbox = vbYes Then
        frmAddMultipleFiles.Show 1
        If lTotalFiles = 0 Then
            ShowMessage "No Files have been Selected: Compile aborted.", Err.Number
            Exit Sub
        End If
    Else
        ShowMessage "No Files have been Selected: Compile aborted.", Err.Number
        Exit Sub
    End If
End If
If Len(txtOutputDirectory.Text) = 0 Then
    mbox = MsgBox("The output directory is not set, would you like to set it now?", vbYesNo + vbQuestion)
    If mbox = vbYes Then
        cmdBrowse_Click
        DoEvents
        If Len(txtOutputDirectory.Text) = 0 Then
            ShowMessage "The Output Directory is not set: Compile aborted.", Err.Number
            Exit Sub
        End If
    Else
        ShowMessage "The Output Directory is not set: Compile aborted.", Err.Number
        Exit Sub
    End If
Else
    If DoesFolderExist(txtOutputDirectory.Text) = False Then
        ShowMessage "The output directory does not exist, compile aborted!", Err.Number
        Exit Sub
    End If
End If
If Len(txtProjectName.Text) = 0 Then
    mbox = MsgBox("Project name has not been set. Would you like to set it now?", vbYesNo + vbQuestion)
    If mbox = vbYes Then
        txtProjectName.Text = InputBox("Enter Project Name:")
        If Len(txtProjectName.Text) = 0 Then
            ShowMessage "The project name has not been set; Compile aborted", Err.Number
            Exit Sub
        End If
    Else
        ShowMessage "The project name has not been set; Compile aborted", Err.Number
        Exit Sub
    End If
End If
Shimmy:
If Len(txtOutputFile.Text) = 0 Then
    mbox = MsgBox("The output file is not set would you like to set it now?", vbYesNo + vbQuestion)
    If mbox = vbYes Then
        'txtOutputFile.Text = GetFileTitle(Trim(SaveDialog(Me, "Executable Files (*.exe)|*.exe|All Files (*.*)|*.*|", "Save EXE File As ...", CurDir)))
        txtOutputFile.Text = Trim(InputBox("Enter output file name"))
        If Len(txtOutputFile.Text) = 0 Then
            ShowMessage "The Output File is not set: Compile Aborted.", Err.Number
            Exit Sub
        End If
        If LCase(Right(txtOutputFile.Text, 4)) <> ".exe" Then txtOutputFile.Text = txtOutputFile.Text & ".exe"
    Else
        ShowMessage "The Output File is not set: Compile Aborted.", Err.Number
        Exit Sub
    End If
End If
If DoesFileExist(txtOutputFile.Text) = True Then
    mbox = MsgBox("The file '" & txtOutputFile.Text & "' already exists in '" & txtOutputDirectory.Text & "'. Would you like to overwrite it?", vbYesNo + vbQuestion)
    If mbox = vbYes Then
        Kill txtOutputDirectory.Text & txtOutputFile.Text
    Else
        txtOutputFile.Text = ""
        GoTo Shimmy
    End If
End If
If DoesFileExist(txtOutputDirectory.Text & txtOutputDirectory.Text) = True Then Exit Sub
If Len(txtOutputDirectory.Text) <> 0 And Right(UCase(txtOutputDirectory.Text), 1) <> "\" Then txtOutputDirectory.Text = txtOutputDirectory.Text & "\"
If Len(txtOutputFile.Text) <> 0 And Right(LCase(txtOutputFile.Text), 4) <> ".exe" Then txtOutputFile.Text = txtOutputFile.Text & ".exe"
mbox = MsgBox("The archive '" & txtOutputFile.Text & "' is ready to be compiled in '" & txtOutputDirectory.Text & "'. Continue?", vbYesNo + vbQuestion)
If mbox = vbNo Then Exit Sub
If lFirstRun = False Then
    MsgBox "Critical Error, haulting program", vbExclamation
    Unload Me
    Exit Sub
Else
    lFirstRun = False
End If
txtPassword.Enabled = False
txtOutputDirectory.Enabled = False
txtOutputFile.Enabled = False
cmdDel.Enabled = False
cmdAdd.Enabled = False
cmdBrowse.Enabled = False
cmdClear.Enabled = False
txtAuthor.Enabled = False
txtLicenseAgreement.Enabled = False
txtProjectName.Enabled = False
cmdCreateArchive.Enabled = False
lPropBag.WriteProperty "FileCount", lTotalFiles
lOutPutFileName = txtOutputDirectory.Text & txtOutputFile.Text
lTempFileHuffman = txtOutputDirectory.Text & "NX_HUFF.TMP"
lTempFileLZSS = txtOutputDirectory.Text & "NX_LZSS.TMP"
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
    StorePropBagSettings
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
msg = "Make archive operation completed" & vbCrLf & vbCrLf & "New Size: " & Format(Outputs, "###,###,###") & " bytes" & vbCrLf & "Difference: " & Format(savings, "###,###,###") & " bytes" & vbCrLf & vbCrLf & "Would you like to test this distribution now? To test click 'Yes', To exit, click 'No'."
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
    MsgBox Err.Description
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
    ShowMessage Err.Description, Err.Number
    Err.Clear
End If
End Function

Private Sub UpDateLast()
On Local Error Resume Next
Dim l As Long
l = FileLen(lOutPutFileName)
ArchSizeLbl.Caption = l
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
    ShowMessage Err.Description, Err.Number
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

Private Sub Form_Load()
On Local Error Resume Next
Dim i As Integer
lEncryptionPassword = "820huiewfkj4hgru32o0ewyuwe"
Set lHuffman = New clsHuffman
lTotalFileSize = 0: lTotalFiles = 0
frmMain.Caption = App.Title & " (" & lTotalFiles & " Files - Size: " & lTotalFileSize & ")"
lTempFileName = App.Path & "\NSSA.EXE"
lFirstRun = True
txtOutputDirectory.Text = App.Path & "\"
Me.Width = ReadINI(App.Path & "\settings.ini", "Main", "Width", Me.Width)
Me.Height = ReadINI(App.Path & "\settings.ini", "Main", "Height", Me.Height)
Me.Top = ReadINI(App.Path & "\settings.ini", "Main", "Top", Me.Top)
Me.Left = ReadINI(App.Path & "\settings.ini", "Main", "Left", Me.Left)
txtAuthor.Text = ReadINI(App.Title & "\settings.ini", "Settings", "Author", "")
End Sub

Private Sub Form_Resize()
On Local Error Resume Next
cmdAbout.Top = Me.ScaleHeight - 460
cmdAbout.Left = Me.ScaleWidth - 3100
cmdExit.Top = Me.ScaleHeight - 460
cmdExit.Left = Me.ScaleWidth - 1570
nTab.Width = Me.ScaleWidth
nTab.Height = Me.ScaleHeight - 900
Shape1.Width = Me.ScaleWidth - 2960
Select Case nTab.ActiveTab
Case 0
    'cmdAdd.Top = Me.ScaleHeight - 1900
    cmdAdd.Top = Me.ScaleHeight - 1400
    cmdDel.Top = Me.ScaleHeight - 1400
    cmdClear.Top = Me.ScaleHeight - 1400
    cmdSetDestination.Top = Me.ScaleHeight - 1400
    cmdSelectAll.Top = Me.ScaleHeight - 1400
    cmdSelectNone.Top = Me.ScaleHeight - 1400
    lvwFiles.Height = Me.ScaleHeight - 1950
    lvwFiles.Width = Me.ScaleWidth - 250
Case 1
    txtProjectName.Width = Me.ScaleWidth - 3000
    txtOutputDirectory.Width = Me.ScaleWidth - 3000
    txtOutputFile.Width = Me.ScaleWidth - 3000
    txtAuthor.Width = Me.ScaleWidth - 3000
    txtLicenseAgreement.Width = Me.ScaleWidth - 3000
    txtPassword.Width = Me.ScaleWidth - 3000
    cmdBrowse.Left = Me.ScaleWidth - cmdBrowse.Width - 100
Case 2
    ctlTotalProgress.Top = Me.ScaleHeight - 1750
    ctlTotalProgress.Width = Me.ScaleWidth - 280
    prgMakeProgress.Top = Me.ScaleHeight - 1350
    prgMakeProgress.Width = Me.ScaleWidth - 280
    cmdCreateArchive.Top = Me.ScaleHeight - 2250
    cmdCreateArchive.Left = Me.ScaleWidth - 1600
    cmdBurnISO.Left = Me.ScaleWidth - 1600
    cmdBurnISO.Top = Me.ScaleHeight - 2360
    cmdRipToIso.Left = Me.ScaleWidth - 1600
    cmdRipToIso.Top = Me.ScaleHeight - 2800
End Select
End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, y As Single)

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Local Error Resume Next
WriteINI App.Path & "\settings.ini", "Main", "Width", Me.Width
WriteINI App.Path & "\settings.ini", "Main", "Height", Me.Height
WriteINI App.Path & "\settings.ini", "Main", "Top", Me.Top
WriteINI App.Path & "\settings.ini", "Main", "Left", Me.Left
WriteINI App.Path & "\settings.ini", "Settings", "Author", txtAuthor.Text
WriteINI App.Path & "\settings.ini", "Settings", "License", txtLicenseAgreement.Text
Unload frmMessages
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

Private Sub lvwFiles_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
On Local Error GoTo ErrHandler
If Len(lvwFiles.SelectedItem.Text) <> 0 And Button = 2 Then
    PopupMenu mnuListViewMenu
End If
Exit Sub
ErrHandler:
    If Err.Number <> 0 Then
        Select Case Err.Number
        Case 0
        Case 91
            Err.Clear
        Case Else
            ShowMessage Err.Description, Err.Number
        End Select
    End If
End Sub

Private Sub lvwFiles_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)
On Local Error GoTo ErrHandler
Dim i As Integer, c As Integer
c = Data.Files.count
For i = 1 To c
    AddFileToListView Data.Files(i), "\"
Next i
Exit Sub
ErrHandler:
    ShowMessage Err.Description, Err.Number
    'MsgBox Err.Description, vbExclamation
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
    ShowMessage Err.Description, Err.Number
    Err.Clear
End Sub

Private Sub nTab_Click()
On Local Error Resume Next
Form_Resize
If Err.Number <> 0 Then
    ShowMessage Err.Description, Err.Number
    Err.Clear
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Local Error Resume Next
Dim msg As String, msg2 As String, mbox As VbMsgBoxResult, msg3 As String, c As Integer, i As Integer
Select Case Button.Index
Case 1
    lvwFiles.ListItems.Clear
Case 2
    msg2 = OpenDialog(Me, "NXT Files (*.nxt)|*.nxt|", App.Title, CurDir)
    If Len(msg2) <> 0 And DoesFileExist(msg2) = True Then
        txtProjectName.Text = ReadINI(msg2, "Settings", "ProjectName", "")
        txtAuthor.Text = ReadINI(msg2, "Settings", "Author", "")
        txtOutputDirectory.Text = ReadINI(msg2, "Settings", "OutputDirectory", "")
        txtOutputFile.Text = ReadINI(msg2, "Settings", "OutputFile", "")
        txtPassword.Text = ReadINI(msg2, "Settings", "Password", "")
        txtLicenseAgreement = ReadINI(msg2, "Settings", "LicenseAgreement", "")
        c = CInt(ReadINI(msg2, "Settings", "FileCount", 0))
        For i = 1 To c
            AddFileToListView Trim(ReadINI(msg2, Trim(CStr(i)), "File", "")), "\"
        Next i
    End If
Case 3
    msg2 = SelectDirectory(Me, "Select Output Dir")
    If Len(Trim(msg2)) <> 0 Then
        msg3 = InputBox("Enter Filename: ", App.Title, lProjectFileName & ".NXT")
        If Len(Trim(msg3)) <> 0 Then
            msg2 = msg2 & msg3
            If LCase(Right(msg3, 4)) <> ".nxt" Then msg3 = msg3 & ".nxt"
            WriteINI msg2, "Settings", "ProjectName", txtProjectName.Text
            WriteINI msg2, "Settings", "Author", txtAuthor.Text
            WriteINI msg2, "Settings", "OutputDirectory", txtOutputDirectory.Text
            WriteINI msg2, "Settings", "OutputFile", txtOutputFile.Text
            WriteINI msg2, "Settings", "Password", txtPassword.Text
            WriteINI msg2, "Settings", "LicenseAgreement", txtLicenseAgreement.Text
            'c = CInt(ReadINI(msg2, "Settings", "FileCount", 0))
            c = lvwFiles.ListItems.count
            WriteINI msg2, "Settings", "FileCount", lvwFiles.ListItems.count
            For i = 1 To c
                WriteINI msg2, Trim(CStr(i)), "File", lvwFiles.ListItems(i).Text
                'AddFileToListView Trim(ReadINI(msg2, Trim(CStr(i)), "File", "")), "\"
            Next i
        End If
    Else
        MsgBox "Dir empty"
    End If
Case 4
    frmAddMultipleFiles.Show
Case 5
    mbox = MsgBox("Are you sure, clear all files?", vbYesNo + vbQuestion)
    If mbox = vbYes Then
        lvwFiles.ListItems.Clear
    End If
Case 6
    msg = InputBox("New Password: ")
    If Len(msg) <> 0 Then txtPassword.Text = msg
Case 7
    cmdCreateArchive_Click
End Select
End Sub
