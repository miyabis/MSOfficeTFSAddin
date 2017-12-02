
Public Interface IDocument

    Function ActiveDocument() As Object

    ReadOnly Property Action As IAction

    Sub Close(Optional ByVal saveChanges As Object = Nothing)

    ReadOnly Property FullName As String

    ReadOnly Property Name As String

    Sub Open(ByVal filename As String)

    Property Saved() As Boolean

    Sub Save()

    Sub SaveAs(ByVal filename As String)

    Sub CompareSideBySideWith(ByVal name As String)

End Interface
