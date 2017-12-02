<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class OutputTaskPaneControl
	Inherits System.Windows.Forms.UserControl

	'UserControl はコンポーネント一覧をクリーンアップするために dispose をオーバーライドします。
	<System.Diagnostics.DebuggerNonUserCode()> _
	Protected Overrides Sub Dispose(ByVal disposing As Boolean)
		Try
			If disposing AndAlso components IsNot Nothing Then
				components.Dispose()
			End If
		Finally
			MyBase.Dispose(disposing)
		End Try
	End Sub

	'Windows フォーム デザイナーで必要です。
	Private components As System.ComponentModel.IContainer

	'メモ: 以下のプロシージャは Windows フォーム デザイナーで必要です。
	'Windows フォーム デザイナーを使用して変更できます。  
	'コード エディターを使って変更しないでください。
	<System.Diagnostics.DebuggerStepThrough()> _
	Private Sub InitializeComponent()
		Me.txtOutput = New System.Windows.Forms.RichTextBox()
		Me.SuspendLayout()
		'
		'txtOutput
		'
		Me.txtOutput.Dock = System.Windows.Forms.DockStyle.Fill
		Me.txtOutput.Location = New System.Drawing.Point(0, 0)
		Me.txtOutput.Name = "txtOutput"
		Me.txtOutput.Size = New System.Drawing.Size(150, 150)
		Me.txtOutput.TabIndex = 0
		Me.txtOutput.Text = ""
		'
		'OutputTaskPaneControl
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.Controls.Add(Me.txtOutput)
		Me.Name = "OutputTaskPaneControl"
		Me.ResumeLayout(False)

	End Sub
	Friend WithEvents txtOutput As System.Windows.Forms.RichTextBox

End Class
