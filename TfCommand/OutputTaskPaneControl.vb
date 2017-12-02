Public Class OutputTaskPaneControl

	Public Sub AddMessage(ByVal value As String)
		Me.txtOutput.AppendText(value)
		Me.txtOutput.AppendText(vbCrLf)
	End Sub

End Class
