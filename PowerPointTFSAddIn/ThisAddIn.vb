
Imports Microsoft.Office.Interop.PowerPoint

Public Class ThisAddIn
    Implements IDocument

    Private _action As IAction

    Private _open As Boolean

#Region " Handles "

#Region " ThisAddIn "

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        Globals.Ribbons.TfsRibbon.Disable()
        _action = New Action(Globals.Ribbons.TfsRibbon, Me)
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

#End Region
#Region " Application "

    Private Sub Application_WindowActivate(Pres As PowerPoint.Presentation, Wn As PowerPoint.DocumentWindow) Handles Application.WindowActivate
        Diagnostics.Debug.Print("Application_WindowActivate:{0}", Wn.Presentation.Name)
        If _open Then
            _open = False
            Return
        End If
        SetEnabled(Wn.Presentation, False)
    End Sub

    Private Sub Application_PresentationOpen(Pres As PowerPoint.Presentation) Handles Application.PresentationOpen
        TF.CreateOutputTaskPane(Pres)
        SetEnabled(Pres, True)
        _open = True
    End Sub

    Private Sub Application_PresentationSave(Pres As PowerPoint.Presentation) Handles Application.PresentationSave
        SetEnabled(Pres, True)
    End Sub

    Private Sub Application_PresentationClose(Pres As PowerPoint.Presentation) Handles Application.PresentationClose
        TF.CloseOutputTaskPane(Pres)
        Globals.Ribbons.TfsRibbon.RemoveEnables(Pres)
    End Sub

#End Region

#End Region

#Region " Implements "

    Public ReadOnly Property Action As IAction Implements IDocument.Action
        Get
            Return _action
        End Get
    End Property

    Public Function ActiveDocument() As Object Implements IDocument.ActiveDocument
        Return Globals.ThisAddIn.Application.ActivePresentation
    End Function

    Public Sub Close(Optional saveChanges As Object = Nothing) Implements IDocument.Close
        Globals.ThisAddIn.Application.ActivePresentation.Close()
    End Sub

    Public Sub CompareSideBySideWith(name As String) Implements IDocument.CompareSideBySideWith
        Globals.ThisAddIn.Application.Windows.CompareSideBySideWith(name)
        Globals.ThisAddIn.Application.Windows.SyncScrollingSideBySide = True
        Globals.ThisAddIn.Application.Windows.Arrange(PowerPoint.PpArrangeStyle.ppArrangeTiled)
    End Sub

    Public ReadOnly Property FullName As String Implements IDocument.FullName
        Get
            Dim filename As String = TF.GetLocalPath(Globals.ThisAddIn.Application.ActivePresentation.FullName)
            Return filename
        End Get
    End Property

    Public ReadOnly Property Name As String Implements IDocument.Name
        Get
            Return Globals.ThisAddIn.Application.ActivePresentation.Name
        End Get
    End Property

    Public Sub Open(filename As String) Implements IDocument.Open
        Globals.ThisAddIn.Application.Presentations.Open(filename)
    End Sub

    Public Sub Save() Implements IDocument.Save
        Globals.ThisAddIn.Application.ActivePresentation.Save()
    End Sub

    Public Sub SaveAs(filename As String) Implements IDocument.SaveAs
        Globals.ThisAddIn.Application.ActivePresentation.SaveAs(filename)
    End Sub

    Public Property Saved As Boolean Implements IDocument.Saved
        Get
            Return Globals.ThisAddIn.Application.ActivePresentation.Saved
        End Get
        Set(value As Boolean)
            Globals.ThisAddIn.Application.ActivePresentation.Saved = value
        End Set
    End Property

#End Region

#Region " Method "

    Private Sub SetEnabled(Wb As PowerPoint.Presentation, ByVal force As Boolean)
        If Wb Is Nothing Then
            Return
        End If

        Globals.Ribbons.TfsRibbon.SetEnabled(Wb, force, True)
    End Sub

    Private Function TF() As TfExe
        If Globals.Ribbons.TfsRibbon.Tf Is Nothing Then
            Globals.Ribbons.TfsRibbon.Init()
        End If
        Return Globals.Ribbons.TfsRibbon.Tf
    End Function

#End Region

End Class
