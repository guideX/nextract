VERSION 5.00
Begin VB.Form frmPassword 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Enter Archive Password"
   ClientHeight    =   1380
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4185
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   4185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   600
      Width           =   3975
   End
   Begin NSSA.ctlButton cmdOK 
      Default         =   -1  'True
      Height          =   345
      Left            =   1800
      TabIndex        =   1
      Top             =   960
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
      Icon            =   "frmPassword1.frx":0000
      Style           =   9
      Caption         =   "OK"
      iNonThemeStyle  =   0
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   1
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   0
      RoundedBordersByTheme=   0   'False
   End
   Begin NSSA.ctlButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   345
      Left            =   3000
      TabIndex        =   2
      Top             =   960
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
      Icon            =   "frmPassword1.frx":001C
      Style           =   9
      Caption         =   "Cancel"
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
      Caption         =   "This archive is password protected. Please enter a password to continue."
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
On Local Error Resume Next
MsgBox "Terminating Archive"
End
End Sub

Private Sub cmdOK_Click()
On Local Error Resume Next
If Trim(LCase(txtPassword.Text)) = Trim(LCase(frmMain.ReturnPassword)) Then
    MsgBox "Password confirmed", vbInformation
    Unload Me
Else
    MsgBox "Password does not match!", vbExclamation
End If
End Sub
