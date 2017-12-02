
Imports System.Windows.Forms
Imports Microsoft.Vbe.Interop

Public Class Vba

    Public Sub Export(ByVal document As IDocument, ByVal vbp As VBProject, ByVal tf As TfExe)
        If Not document.Saved Then
            Dim rc As DialogResult
            rc = MessageBox.Show(TfCommand.My.Resources.Messages.Q002,
                                 TfCommand.My.Resources.Messages.MessageBoxTitle,
                                 MessageBoxButtons.YesNoCancel,
                                 MessageBoxIcon.Question,
                                 MessageBoxDefaultButton.Button3)
            If rc = DialogResult.Cancel Then
                Return
            End If
            If rc = DialogResult.Yes Then
                document.Save()
            End If
        End If
        Dim names() As String = document.Name.Split(".")
        If names.Count.Equals(1) Then
            Return
        End If

        Dim path As String
        path = IO.Path.Combine(document.FullName.Replace(document.Name, String.Empty), document.Name & ".vba")

        Using dlg As New System.Windows.Forms.FolderBrowserDialog
            dlg.RootFolder = Environment.SpecialFolder.Desktop
            dlg.SelectedPath = IO.Path.GetDirectoryName(path)
            dlg.ShowNewFolderButton = True
            If dlg.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then
                Return
            End If
            path = dlg.SelectedPath
        End Using

        Dim addflag As Boolean = tf.IsAdd(document.Name)

        For Each component As VBComponent In vbp.VBComponents
            Dim fileExt As String
            Dim fullname As String
            fileExt = _getExt(component)
            If fileExt = String.Empty Then
                Continue For
            End If

            fullname = _export(component, path, fileExt)
            If addflag Then
                Continue For
            End If

            tf.Add(fullname)
            If Not fullname.EndsWith(".frm") Then
                Continue For
            End If

            tf.Add(fullname.Replace(".frm", ".frx"))
        Next
    End Sub

    Private Function _getExt(ByVal component As VBComponent) As String
        Select Case component.Type
            Case vbext_ComponentType.vbext_ct_StdModule
                Return "bas"
            Case vbext_ComponentType.vbext_ct_ClassModule, vbext_ComponentType.vbext_ct_Document
                Return "cls"
            Case vbext_ComponentType.vbext_ct_MSForm
                Return "frm"
            Case Else
                Return "bas"
        End Select
    End Function

    Private Function _export(ByVal component As VBComponent, ByVal outPath As String, ByVal ext As String) As String
        Dim dinfo As New IO.DirectoryInfo(outPath)
        Dim fullname As String = IO.Path.Combine(dinfo.FullName, component.Name & "." & ext)

        'If component.CodeModule.CountOfLines.Equals(0) Then
        '    If IO.File.Exists(fullname) Then
        '        IO.File.Delete(fullname)
        '        If fullname.EndsWith(".frm") Then
        '            IO.File.Delete(fullname.Replace("frm", "frx"))
        '        End If
        '    End If
        '    Return
        'End If

        If Not dinfo.Exists Then
            dinfo.Create()
        End If

        component.Export(fullname)
        '_cnvEnc(fullname)

        Return fullname
    End Function

    Private Sub _cnvEnc(ByVal filename As String)
        Dim s As String
        Using sr As New System.IO.StreamReader(filename, System.Text.Encoding.GetEncoding("shift_jis"))
            s = sr.ReadToEnd()
        End Using
        If filename.EndsWith(".frx") Then
            Return
        End If
        Using sw As New System.IO.StreamWriter(filename, False)
            sw.Write(s)
        End Using
    End Sub

End Class
