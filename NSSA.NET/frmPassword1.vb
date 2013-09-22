Option Strict Off
Option Explicit On
Friend Class frmPassword
	Inherits System.Windows.Forms.Form
	
	Private Sub cmdCancel_Click(ByVal Sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
		On Error Resume Next
		MsgBox("Terminating Archive")
		End
	End Sub
	
	Private Sub cmdOK_Click(ByVal Sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click
		On Error Resume Next
		If Trim(LCase(txtPassword.Text)) = Trim(LCase(frmMain.ReturnPassword)) Then
			MsgBox("Password confirmed", MsgBoxStyle.Information)
			Me.Close()
		Else
			MsgBox("Password does not match!", MsgBoxStyle.Exclamation)
		End If
	End Sub
End Class