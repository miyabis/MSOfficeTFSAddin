
Imports System.IO
Imports System.Diagnostics
Imports System.Text
Imports Microsoft.Win32

''' <summary>
''' TortoiseProcExe コマンド実行
''' </summary>
''' <remarks></remarks>
Public Class TortoiseProcExe

#Region " Declare "

    Private _exePath As String
    Private _exe As Process
    Private _exeInfo As ProcessStartInfo

    Private _exitCode As Integer
    Private _standardOutput As String
    Private _standardError As String

#End Region

#Region " コンストラクタ "

    ''' <summary>
    ''' デフォルトコンストラクタ
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()
        _init()
    End Sub

#End Region
#Region " Property "

    ''' <summary>
    ''' 実行可能かどうか
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property IsExecute As Boolean
        Get
            Return Not String.IsNullOrEmpty(_exePath)
        End Get
    End Property

#End Region
#Region " Method "

    ''' <summary>
    ''' 比較実行
    ''' </summary>
    ''' <param name="filename"></param>
    ''' <param name="filename2"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Diff(ByVal filename As String, ByVal filename2 As String) As Boolean
        Const C_CMD As String = "/command:diff /path:""{0}"" /path2:""{1}"""
        Dim args As StringBuilder = New StringBuilder
        args.AppendFormat(C_CMD, filename, filename2)
        Return _commandExecute(args.ToString)
    End Function

    ''' <summary>
    ''' 初期化
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub _init()
        Const C_VALUE As String = "ProcPath"
        Dim key As RegistryKey = Nothing
        key = Registry.LocalMachine.OpenSubKey("SOFTWARE\TortoiseSVN")
        If key Is Nothing Then
            key = Registry.LocalMachine.OpenSubKey("SOFTWARE\TortoiseGit")
            If key Is Nothing Then
                Return
            End If
        End If
        _exePath = key.GetValue(C_VALUE)

        _exeInfo = New ProcessStartInfo()
        _exeInfo.LoadUserProfile = True
        _exeInfo.UseShellExecute = False
        _exeInfo.CreateNoWindow = True
        _exeInfo.RedirectStandardError = True
        _exeInfo.RedirectStandardOutput = True
        _exeInfo.FileName = _exePath
    End Sub

    ''' <summary>
    ''' コマンド実行
    ''' </summary>
    ''' <param name="args"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function _commandExecute(ByVal args As String) As Boolean
        If Not Me.IsExecute Then
            Return False
        End If

        _exeInfo.Arguments = args

        _exe = Process.Start(_exeInfo)
        _exe.WaitForExit()

        _standardOutput = _exe.StandardOutput.ReadToEnd()
        _standardError = _exe.StandardError.ReadToEnd()
        _exitCode = _exe.ExitCode
        If Not _exitCode.Equals(0) Then
            Return False
        End If
        Return True
    End Function

#End Region

End Class
