
Public Class ActionContext

    Private _documnet As IDocument
    Private _name As String

    Public Sub New(ByVal documnet As IDocument)
        _documnet = documnet
        _name = documnet.Name
        Me.FullName = documnet.FullName
    End Sub

    Public ReadOnly Property Documnet As IDocument
        Get
            Return _documnet
        End Get
    End Property

    Public ReadOnly Property Name As String
        Get
            Return _name
        End Get
    End Property

    Public Property FullName As String

    Public Property StandardOutput As String

    Public Property StandardError As String

    Public Property OutputActive As Boolean

    Public Property StatusBar As Object

End Class
