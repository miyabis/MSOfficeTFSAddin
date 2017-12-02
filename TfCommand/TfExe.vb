
Imports System.IO
Imports System.Diagnostics
Imports System.Text
Imports System.Windows.Forms
Imports Microsoft.Win32

''' <summary>
''' tf.exe コマンド実行
''' </summary>
''' <remarks></remarks>
Public Class TfExe

#Region " Declare "

    Private _vsversion As Integer

    Private tfExePath As String
    Private tfExe As Process
    Private tfExeInfo As ProcessStartInfo

    Private _exitCode As Integer
    Private _standardOutput As String
    Private _standardError As String

    Private _beforeCommandInfo As Boolean

    Private _customTaskPane As Microsoft.Office.Tools.CustomTaskPaneCollection
    Private _customTaskPanes As IDictionary(Of Object, Microsoft.Office.Tools.CustomTaskPane)

    Private _outputTaskPaneControl As OutputTaskPaneControl

#End Region

#Region " コンストラクタ "

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <param name="customTaskPane"></param>
    ''' <remarks></remarks>
    Public Sub New(ByVal customTaskPane As Microsoft.Office.Tools.CustomTaskPaneCollection)
        _customTaskPane = customTaskPane
        _init()

        _customTaskPanes = New Dictionary(Of Object, Microsoft.Office.Tools.CustomTaskPane)
    End Sub

#End Region
#Region " Property "

    Public Property InstallDir As String

    Public ReadOnly Property VSVersion As Integer
        Get
            Return _vsversion
        End Get
    End Property

    Public ReadOnly Property IsExecute As Boolean
        Get
            _beforeCommandInfo = False
            Return Not String.IsNullOrEmpty(tfExePath)
        End Get
    End Property

    Public ReadOnly Property StandardOutput As String
        Get
            Return _standardOutput
        End Get
    End Property

    Public ReadOnly Property StandardError As String
        Get
            Return _standardError
        End Get
    End Property

    Public ReadOnly Property IsAdd(ByVal filename As String) As Boolean
        Get
            Const C_CMD As String = "status ""{0}"""
            Dim args As StringBuilder = New StringBuilder
            args.AppendFormat(C_CMD, filename)
            If Not _commandExecute(args.ToString) Then
                Return False
            End If

            Dim r As New RegularExpressions.Regex("(\d) .*(\d) .*$", RegularExpressions.RegexOptions.IgnoreCase)
            Dim mc As RegularExpressions.MatchCollection = r.Matches(Me.StandardOutput)
            Return mc.Count > 0
        End Get
    End Property

    Public ReadOnly Property IsModify(ByVal filename As String) As Boolean
        Get
            Const C_CMD As String = "status ""{0}"""
            Dim args As StringBuilder = New StringBuilder
            args.AppendFormat(C_CMD, filename)
            If Not _commandExecute(args.ToString) Then
                Return False
            End If

            Dim r As New RegularExpressions.Regex("(\d) .*$", RegularExpressions.RegexOptions.IgnoreCase)
            Dim mc As RegularExpressions.MatchCollection = r.Matches(Me.StandardOutput)
            Return mc.Count > 0
        End Get
    End Property

    Public ReadOnly Property IsDiff(ByVal filename As String) As Boolean
        Get
            If Not Me.Difference(filename) Then
                Return False
            End If
            filename = Path.GetFileName(filename)
            Return ((Me.StandardOutput.Length - Me.StandardOutput.Replace(filename, "").Length) \ filename.Length) = 2
        End Get
    End Property

    Public Property WorkingDirectory As String
        Get
            Return tfExeInfo.WorkingDirectory
        End Get
        Set(value As String)
            tfExeInfo.WorkingDirectory = value
        End Set
    End Property

#End Region
#Region " Method "

    Public Sub ShowOutputTaskPane(ByVal doc As Object)
        If Not _customTaskPanes.ContainsKey(doc) Then
            CreateOutputTaskPane(doc)
        End If

        Dim pane As Microsoft.Office.Tools.CustomTaskPane
        pane = DirectCast(_customTaskPanes.Item(doc), Microsoft.Office.Tools.CustomTaskPane)
        If pane.Visible Then
            Return
        End If
        pane.Visible = True
    End Sub

    Public Sub CreateOutputTaskPane(ByVal doc As Object)
        If _customTaskPanes.ContainsKey(doc) Then
            Return
        End If

        Dim outputTaskPaneControl As OutputTaskPaneControl
        Dim pane As Microsoft.Office.Tools.CustomTaskPane

        outputTaskPaneControl = New OutputTaskPaneControl
        pane = _customTaskPane.Add(outputTaskPaneControl, "TfsOfficeAddIn Output")
        pane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionBottom
        pane.Height = 150
        pane.Visible = False

        _customTaskPanes.Add(doc, pane)
    End Sub

    Public Sub CloseOutputTaskPane(ByVal doc As Object)
        If doc Is Nothing Then
            Return
        End If
        If Not _customTaskPanes.ContainsKey(doc) Then
            Return
        End If

        Dim pane As Microsoft.Office.Tools.CustomTaskPane
        pane = DirectCast(_customTaskPanes.Item(doc), Microsoft.Office.Tools.CustomTaskPane)

        _customTaskPanes.Remove(doc)
        _customTaskPane.Remove(pane)
        pane.Dispose()
    End Sub

    Public Sub AddMessage(ByVal doc As Object, ByVal value As String)
        If Not _customTaskPanes.ContainsKey(doc) Then
            CreateOutputTaskPane(doc)
        End If
        Dim pane As Microsoft.Office.Tools.CustomTaskPane
        pane = DirectCast(_customTaskPanes.Item(doc), Microsoft.Office.Tools.CustomTaskPane)
        Dim ctrl As OutputTaskPaneControl
        ctrl = DirectCast(pane.Control, OutputTaskPaneControl)
        ctrl.AddMessage(value)
        If pane.Visible Then
            Return
        End If
        pane.Visible = True
    End Sub

    Public Sub OpenExplorer(ByVal filename As String)
        Const C_ARGS As String = "/select,""{0}"""
        Process.Start("explorer", String.Format(C_ARGS, filename))
    End Sub

    Public Sub ExplorerExecute(ByVal filename As String)
        Const C_ARGS As String = """{0}"""
        Process.Start("explorer", String.Format(C_ARGS, filename))
    End Sub

    Public Sub OpenVisualStudio()
        Const C_DEV As String = "devenv.exe"
        Process.Start(Path.Combine(Me.InstallDir, C_DEV))
    End Sub

    Public Function Workfold() As Boolean
        Const C_CMD As String = "workfold"
        Dim args As StringBuilder = New StringBuilder
        args.Append(C_CMD)
        Return _commandExecute(args.ToString)
    End Function

    Public Function Workfold(ByVal filename As String) As Boolean
        Const C_CMD As String = "workfold ""{0}"""
        Dim args As StringBuilder = New StringBuilder
        args.AppendFormat(C_CMD, filename)
        Return _commandExecute(args.ToString)
    End Function

    Public Function GetWebAccess() As String
        Const C_CMD As String = "workfold"
        Dim args As StringBuilder = New StringBuilder
        args.Append(C_CMD)
        If Not _commandExecute(args.ToString) Then
            Return String.Empty
        End If

        Dim url As New StringBuilder
        Dim r As New RegularExpressions.Regex("https?://[\w/:%#\$&\?\(\)~\.=\+\-]+", RegularExpressions.RegexOptions.IgnoreCase)
        Dim mc As RegularExpressions.MatchCollection = r.Matches(Me.StandardOutput)
        If mc.Count.Equals(0) Then
            Return String.Empty
        End If
        url.Append(mc(0).Value)

        r = New RegularExpressions.Regex(" \$/(.*): ", RegularExpressions.RegexOptions.IgnoreCase)
        mc = r.Matches(Me.StandardOutput)
        If Not mc.Count.Equals(0) Then
            url.Append("/")
            url.Append(mc(0).Groups(1).Value)
        End If

        Return url.ToString
    End Function

    Public Function Status(ByVal filename As String) As Boolean
        Const C_CMD As String = "status ""{0}"""
        Dim args As StringBuilder = New StringBuilder
        args.AppendFormat(C_CMD, filename)
        If Not _commandExecute(args.ToString) Then
            Return False
        End If
        Return _standardOutput.Contains(filename)
    End Function

    Public Function Status2(ByVal filename As String) As Boolean
        Const C_CMD As String = "status ""{0}"""
        Dim args As StringBuilder = New StringBuilder
        args.AppendFormat(C_CMD, filename)
        If Not _commandExecute(args.ToString) Then
            Return False
        End If
        Return True
    End Function

    Public Function Info(ByVal filename As String) As Boolean
        Const C_CMD10 As String = "properties ""{0}"""  ' 2010
        Const C_CMD11 As String = "info ""{0}"""    ' 2013
        Dim cmd As String
        If Me.VSVersion.Equals(10) Then
            cmd = C_CMD10
        Else
            cmd = C_CMD11
        End If
        Dim args As StringBuilder = New StringBuilder
        args.AppendFormat(cmd, filename)
        If Not _commandExecute(args.ToString) Then
            Return False
        End If
        _beforeCommandInfo = True
        Return _standardOutput.Contains(filename)
    End Function

    Public Function Difference(ByVal filename As String) As Boolean
        Const C_CMD As String = "difference ""{0}"" /format:Brief"
        'Const C_CMD As String = "difference ""{0}"""
        Dim args As StringBuilder = New StringBuilder
        args.AppendFormat(C_CMD, Path.GetFileName(filename))
        Return _commandExecute(args.ToString)
    End Function

    Public Function Add(ByVal filename As String) As Boolean
        Const C_CMD As String = "add ""{0}"""
        Dim args As StringBuilder = New StringBuilder
        args.AppendFormat(C_CMD, filename)
        Return _commandExecute(args.ToString)
    End Function

    Public Function Undo(ByVal filename As String) As Boolean
        Const C_CMD As String = "undo ""{0}"" /noprompt"
        Dim args As StringBuilder = New StringBuilder
        args.AppendFormat(C_CMD, filename)
        Return _commandExecute(args.ToString)
        'args.AppendFormat(C_CMD, Path.GetFileName(filename))
        'Return _commandExecute2(args.ToString, Path.GetDirectoryName(filename))
    End Function

    Public Function CheckIn(ByVal filename As String) As Boolean
        Const C_CMD As String = "checkin ""{0}"""
        Dim args As StringBuilder = New StringBuilder
        args.AppendFormat(C_CMD, Path.GetFileName(filename))
        Return _commandExecute2(args.ToString, Path.GetDirectoryName(filename))
    End Function

    Public Function CheckOut(ByVal filename As String) As Boolean
        Const C_CMD As String = "checkout ""{0}"""
        Dim args As StringBuilder = New StringBuilder
        args.AppendFormat(C_CMD, filename)
        Return _commandExecute(args.ToString)
    End Function

    ''' <summary>
    ''' バージョンの間の競合を解決
    ''' </summary>
    ''' <param name="filename"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Resolve(ByVal filename As String) As Boolean
        Const C_CMD As String = "resolve ""{0}"""
        Dim args As StringBuilder = New StringBuilder
        args.AppendFormat(C_CMD, Path.GetFileName(filename))
        Return _commandExecute2(args.ToString, Path.GetDirectoryName(filename))
    End Function

    Public Function [Get](Optional ByVal filename As String = "") As Boolean
        Const C_CMD As String = "get ""{0}"""
        Const C_CMD2 As String = "get"
        Dim args As StringBuilder = New StringBuilder
        If String.IsNullOrEmpty(filename) Then
            args.Append(C_CMD2)
        Else
            args.AppendFormat(C_CMD, filename)
        End If
        Return _commandExecute(args.ToString)
    End Function

    Public Function History(ByVal filename As String) As Boolean
        Const C_CMD As String = "history ""{0}"""
        Dim args As StringBuilder = New StringBuilder
        args.AppendFormat(C_CMD, Path.GetFileName(filename))
        Return _commandExecute2(args.ToString, Path.GetDirectoryName(filename))
    End Function

    Public Function Shelve(ByVal filename As String) As Boolean
        Const C_CMD As String = "shelve ""{0}"""
        Dim args As StringBuilder = New StringBuilder
        args.AppendFormat(C_CMD, Path.GetFileName(filename))
        Return _commandExecute2(args.ToString, Path.GetDirectoryName(filename))
    End Function

    Public Function Unshelve(ByVal filename As String) As Boolean
        Const C_CMD As String = "unshelve ""{0}"""
        Dim args As StringBuilder = New StringBuilder
        args.AppendFormat(C_CMD, Path.GetFileName(filename))
        Return _commandExecute2(args.ToString, Path.GetDirectoryName(filename))
    End Function

    Public Function Rename(ByVal filename As String, ByVal filenameAs As String) As Boolean
        Const C_CMD As String = "rename ""{0}"" ""{1}"""
        Dim args As StringBuilder = New StringBuilder
        args.AppendFormat(C_CMD, Path.GetFileName(filename), filenameAs)
        'Return _commandExecute2(args.ToString, Path.GetDirectoryName(filename))
        Return _commandExecute(args.ToString)
    End Function

    Private Sub _init()
        Const C_TFEXE As String = "TF.exe"
        Const C_VALUE As String = "InstallDir"
        Dim vers() As Integer = {14, 12, 11, 10}
        Dim key As RegistryKey = Nothing
        For Each ver As Integer In vers
            key = Registry.CurrentUser.OpenSubKey(String.Format("Software\Microsoft\VisualStudio\{0}.0_Config", ver))
            If key IsNot Nothing Then
                _vsversion = ver
                Exit For
            End If
        Next
        If key Is Nothing Then
            Return
        End If
        Me.InstallDir = key.GetValue(C_VALUE)
        tfExePath = Path.Combine(Me.InstallDir, C_TFEXE)

        tfExeInfo = New ProcessStartInfo()
        tfExeInfo.LoadUserProfile = True
        tfExeInfo.UseShellExecute = False
        tfExeInfo.CreateNoWindow = True
        tfExeInfo.RedirectStandardError = True
        tfExeInfo.RedirectStandardOutput = True
        tfExeInfo.FileName = tfExePath
    End Sub

    Private Function _commandExecute(ByVal args As String) As Boolean
        If Not Me.IsExecute Then
            Return False
        End If

        tfExeInfo.Arguments = args

        tfExe = Process.Start(tfExeInfo)
        tfExe.WaitForExit()

        _standardOutput = tfExe.StandardOutput.ReadToEnd()
        _standardError = tfExe.StandardError.ReadToEnd()
        _exitCode = tfExe.ExitCode
        If Not _exitCode.Equals(0) Then
            Return False
        End If
        Return True
    End Function

    Private Function _commandExecute2(ByVal args As String, ByVal workingDirectory As String) As Boolean
        Dim info As New ProcessStartInfo()
        info.FileName = "cmd.exe"
        info.WorkingDirectory = workingDirectory
        info.Arguments = String.Format("/c """"{0}"" {1}""", tfExePath, args)
        'info.Arguments = String.Format("/c ""{0}"" {1}", tfExePath, args)
        info.UseShellExecute = True
        info.WindowStyle = ProcessWindowStyle.Hidden
        Dim exe As Process = Process.Start(info)
        exe.WaitForExit()
        _exitCode = exe.ExitCode
        If Not _exitCode.Equals(0) Then
            Return False
        End If
        Return True
    End Function

#End Region

End Class
