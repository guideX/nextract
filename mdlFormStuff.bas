Attribute VB_Name = "mdlFormStuff"
Option Explicit
Private lMainForm As New frmMain
Private lButtonType As Integer
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Public Function ReturnMainFormWidth() As Long
On Local Error Resume Next
If ReturnShowDialog = True Then ReturnMainFormWidth = lMainForm.Width
End Function

Public Function ReturnMainFormHeight() As Long
On Local Error Resume Next
If ReturnShowDialog = True Then ReturnMainFormHeight = lMainForm.Height
End Function

Public Function ReturnMainFormLeft() As Long
On Local Error Resume Next
If ReturnShowDialog = True Then ReturnMainFormLeft = lMainForm.Left
End Function

Public Function ReturnMainFormTop() As Long
On Local Error Resume Next
If ReturnShowDialog = True Then ReturnMainFormTop = lMainForm.Top
End Function

Public Sub SetMainFormBackground(lBackground As String)
On Local Error Resume Next
Dim i As Integer, w As Boolean
If ReturnSettingsUseThemeColor = True Then
    If ReturnShowDialog = True Then
        Select Case Int(lBackground)
        Case 0
            lBackground = RGB(255, 255, 255)
            w = True
        Case 1
            lBackground = RGB(0, 0, 0)
        Case 2
            lBackground = RGB(0, 0, 127)
        Case 3
            lBackground = RGB(0, 147, 0)
        Case 4
            lBackground = RGB(255, 0, 0)
        Case 5
            lBackground = RGB(127, 0, 0)
        Case 6
            lBackground = RGB(0, 0, 0)
            'lBackground = RGB(156, 0, 156)
        Case 7
            lBackground = RGB(252, 127, 0)
            w = True
        Case 8
            lBackground = RGB(255, 255, 0)
            w = True
        Case 9
            lBackground = RGB(0, 252, 0)
        Case 10
            lBackground = RGB(0, 147, 147)
        Case 11
            lBackground = RGB(0, 255, 255)
            w = True
        Case 12
            lBackground = RGB(0, 0, 252)
        Case 13
            lBackground = RGB(255, 0, 255)
            w = True
        Case 14
            lBackground = RGB(127, 127, 127)
        Case 15
            lBackground = RGB(210, 210, 210)
            w = True
        End Select
        With lMainForm
            .BackColor = lBackground
            .Refresh
        End With
    End If
End If
End Sub

Public Function DoesListBoxItemExist(lText As String, lListBox As ListBox) As Boolean
On Local Error Resume Next
Dim i As Integer
If ReturnShowDialog = True Then
    For i = 0 To lListBox.ListCount
        If LCase(Trim(lText)) = LCase(Trim(lListBox.List(i))) Then
            DoesListBoxItemExist = True
            Exit For
        End If
    Next i
End If
End Function

Public Function FindListBoxIndex(lText As String, lListBox As ListBox) As Integer
On Local Error Resume Next
Dim i As Integer
If ReturnShowDialog = True Then
    For i = 0 To lListBox.ListCount
        If LCase(Trim(lText)) = LCase(Trim(lListBox.List(i))) Then
            FindListBoxIndex = i
            Exit Function
        End If
    Next i
End If
End Function

Public Function GetMainFormButtonType() As Integer
On Local Error Resume Next
If ReturnShowDialog = True Then GetMainFormButtonType = lButtonType
End Function

Public Sub SetMainFormButtonType(lType As Integer)
On Local Error Resume Next
If ReturnShowDialog = True Then lButtonType = lType
End Sub

Public Function ReturnMainFormCaption() As String
On Local Error Resume Next
If ReturnShowDialog = True Then ReturnMainFormCaption = lMainForm.Caption
End Function

Public Sub CloseMainForm()
On Local Error Resume Next
If ReturnShowDialog = True Then Unload lMainForm
End Sub

Public Function ReturnMainForm() As Form
On Local Error Resume Next
If ReturnShowDialog = True Then Set ReturnMainForm = lMainForm
End Function

Public Sub LoadMainForm(lHwnd As Long)
On Local Error GoTo ErrHandler
'If ReturnShowDialog = True Then
Set lMainForm = New frmMain
SetChildHWND CLng(lHwnd)
ShowNonModalForm lMainForm
'Else
'End If
Exit Sub
ErrHandler:
    Err.Clear
End Sub

Public Sub GoURL1(lURL As String)
On Local Error GoTo ErrHandler
lMainForm.ctlBrowser.Navigate lURL
Exit Sub
ErrHandler:
    Err.Clear
End Sub
