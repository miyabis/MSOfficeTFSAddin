
Imports System.IO
Imports System.Diagnostics
Imports System.Windows.Forms
Imports Microsoft.Office.Tools.Ribbon
Imports Microsoft.Win32
Imports MiYABiS

''' <summary>
''' アクション
''' </summary>
''' <remarks></remarks>
Public Class Action
    Implements IAction

    Private _ribbon As TfsRibbon
    Private _document As IDocument

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <param name="ribbon"></param>
    ''' <param name="document"></param>
    ''' <remarks></remarks>
    Public Sub New(ByVal ribbon As TfsRibbon, ByVal document As IDocument)
        _ribbon = ribbon
        _document = document
    End Sub

    ''' <summary>
    ''' 実行
    ''' </summary>
    ''' <param name="method"></param>
    ''' <remarks></remarks>
    Public Sub Execute(method As ExecuteMethod) Implements IAction.Execute
        Dim context As New ActionContext(_document)

        Try
            ' コマンド実行
            method(context)
        Catch ex As Exception
            context.StandardError &= vbCrLf & ex.Message
        End Try

        ' 状態チェック
        _ribbon.SetEnabled(_document.ActiveDocument, True)

        If context.OutputActive Then
            _ribbon.AddMessage(context.StandardOutput)
        End If
        If Not String.IsNullOrEmpty(context.StatusBar) Then
            MessageBox.Show(context.StatusBar,
                            TfCommand.My.Resources.Messages.MessageBoxTitle,
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Exclamation)
        End If

        If String.IsNullOrEmpty(context.StandardError) Then
            Return
        End If

        MessageBox.Show(context.StandardError,
                        TfCommand.My.Resources.Messages.MessageBoxTitle,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Exclamation)
        _ribbon.AddMessage(context.StandardError)
    End Sub

    ''' <summary>
    ''' 実行（閉じて開きなおす）
    ''' </summary>
    ''' <param name="method"></param>
    ''' <remarks></remarks>
    Public Sub ExecuteCloseOpen(method As ExecuteMethod) Implements IAction.ExecuteCloseOpen
        Dim context As New ActionContext(_document)

        ' 保存していないときは保存する？
        Dim rc As DialogResult
        If Not _document.Saved Then
            rc = MessageBox.Show(TfCommand.My.Resources.Messages.Q002,
                                 TfCommand.My.Resources.Messages.MessageBoxTitle,
                                 MessageBoxButtons.YesNoCancel,
                                 MessageBoxIcon.Question,
                                 MessageBoxDefaultButton.Button3)
            If rc = DialogResult.Cancel Then
                Return
            End If
            If rc = DialogResult.Yes Then
                _document.Save()
            End If
        End If

        ' ブックを閉じる
        _document.Close(False)

        Try
            ' コマンド実行
            method(context)
        Catch ex As Exception
            context.StandardError &= vbCrLf & ex.Message
        End Try

        If Not File.Exists(context.FullName) Then
            Dim ofd As New OpenFileDialog()
            ofd.AddExtension = True
            ofd.DefaultExt = Path.GetExtension(context.Name)
            ofd.InitialDirectory = Path.GetDirectoryName(context.FullName)
            ofd.RestoreDirectory = True
            If ofd.ShowDialog() <> DialogResult.OK Then
                Return
            End If
            context.FullName = ofd.FileName
        End If

        If File.Exists(context.FullName) Then
            ' ブックを開きなおす
            _document.Open(context.FullName)
        End If

        If context.OutputActive Then
            _ribbon.AddMessage(context.StandardOutput)
        End If
        If Not String.IsNullOrEmpty(context.StatusBar) Then
            MessageBox.Show(context.StatusBar,
                            TfCommand.My.Resources.Messages.MessageBoxTitle,
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Exclamation)
        End If

        If String.IsNullOrEmpty(context.StandardError) Then
            Return
        End If

        MessageBox.Show(context.StandardError,
                        TfCommand.My.Resources.Messages.MessageBoxTitle,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Exclamation)
        _ribbon.AddMessage(context.StandardError)
    End Sub

End Class
