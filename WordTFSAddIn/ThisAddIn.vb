
Imports System.Threading

Public Class ThisAddIn
    Implements IDocument

    Private _action As IAction

    Private Event DocumentAfterSave(ByVal Doc As Word.Document)
    Private Delegate Sub _saveCheckDelegate()
    Private _saveCheck As _saveCheckDelegate
    Private _saveDone As AsyncCallback
    Private _queueDocs As New System.Collections.Queue

    Private _open As Boolean

#Region " Handles "

#Region " ThisAddIn "

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        Globals.Ribbons.TfsRibbon.Disable()
        _action = New Action(Globals.Ribbons.TfsRibbon, Me)

        _saveCheck = New _saveCheckDelegate(AddressOf _afterSave)
        _saveDone = New System.AsyncCallback(AddressOf _afterSaveDone)
        _saveCheck.BeginInvoke(_saveDone, _saveCheck)
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
    End Sub

#End Region
#Region " Application "

    Private Sub Application_DocumentChange() Handles Application.DocumentChange
        If _open Then
            _open = False
            Return
        End If
        SetEnabled(Me.ActiveDocument, False)
        _open = True
    End Sub

    Private Sub Application_WindowActivate(Doc As Microsoft.Office.Interop.Word.Document, Wn As Microsoft.Office.Interop.Word.Window) Handles Application.WindowActivate
        If _open Then
            _open = False
            Return
        End If
        TF.CreateOutputTaskPane(Doc)
        SetEnabled(Doc, False)
        _open = True
    End Sub

    Private Sub Application_DocumentOpen(Doc As Microsoft.Office.Interop.Word.Document) Handles Application.DocumentOpen
        TF.CreateOutputTaskPane(Doc)
        SetEnabled(Doc, True)
        _open = True
    End Sub

    Private Sub Application_DocumentBeforeClose(Doc As Microsoft.Office.Interop.Word.Document, ByRef Cancel As Boolean) Handles Application.DocumentBeforeClose
        TF.CloseOutputTaskPane(Doc)
        Globals.Ribbons.TfsRibbon.RemoveEnables(Doc)
    End Sub

    Private Sub ThisAddIn_DocumentAfterSave(Doc As Microsoft.Office.Interop.Word.Document) Handles Me.DocumentAfterSave
        SetEnabled(Doc, True)
    End Sub

    Private Sub Application_DocumentBeforeSave(Doc As Microsoft.Office.Interop.Word.Document, ByRef SaveAsUI As Boolean, ByRef Cancel As Boolean) Handles Application.DocumentBeforeSave
        SyncLock Me._queueDocs
            If Not _queueDocs.Contains(Doc) Then
                _queueDocs.Enqueue(Doc)
            End If
        End SyncLock
    End Sub

    Private Sub _afterSave()
        If Me._queueDocs.Count <= 0 Then
            Thread.Sleep(500)
            Return
        End If
        SyncLock Me._queueDocs
            Try
                Dim curDoc As Word.Document = _queueDocs.Peek
                If curDoc.Saved Then
                    RaiseEvent DocumentAfterSave(_queueDocs.Dequeue)
                End If
                curDoc = Nothing
            Catch ex As Exception
                _saveCheck.BeginInvoke(_saveDone, _saveCheck)
            End Try
        End SyncLock
    End Sub
    Private Sub _afterSaveDone(ByVal ar As System.IAsyncResult)
        _saveCheck.BeginInvoke(_saveDone, _saveCheck)
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
        Try
            Return Globals.ThisAddIn.Application.ActiveDocument
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Public Sub Close(Optional saveChanges As Object = Nothing) Implements IDocument.Close
        Globals.ThisAddIn.Application.ActiveDocument.Close(saveChanges)
    End Sub

    Public Sub CompareSideBySideWith(name As String) Implements IDocument.CompareSideBySideWith
        Globals.ThisAddIn.Application.Windows.CompareSideBySideWith(name)
        Globals.ThisAddIn.Application.Windows.SyncScrollingSideBySide = True
        Globals.ThisAddIn.Application.Windows.Arrange(Word.WdArrangeStyle.wdTiled)
    End Sub

    Public ReadOnly Property FullName As String Implements IDocument.FullName
        Get
            Dim filename As String = TF.GetLocalPath(Globals.ThisAddIn.Application.ActiveDocument.FullName)
            Return filename
        End Get
    End Property

    Public ReadOnly Property Name As String Implements IDocument.Name
        Get
            Return Globals.ThisAddIn.Application.ActiveDocument.Name
        End Get
    End Property

    Public Sub Open(filename As String) Implements IDocument.Open
        Globals.ThisAddIn.Application.Documents.Open(filename)
    End Sub

    Public Sub Save() Implements IDocument.Save
        Globals.ThisAddIn.Application.ActiveDocument.Save()
    End Sub

    Public Sub SaveAs(filename As String) Implements IDocument.SaveAs
        Globals.ThisAddIn.Application.ActiveDocument.SaveAs(filename)
    End Sub

    Public Property Saved As Boolean Implements IDocument.Saved
        Get
            Return Globals.ThisAddIn.Application.ActiveDocument.Saved
        End Get
        Set(value As Boolean)
            Globals.ThisAddIn.Application.ActiveDocument.Saved = value
        End Set
    End Property

#End Region

#Region " Method "

    Private Sub SetEnabled(Wb As Microsoft.Office.Interop.Word.Document, ByVal force As Boolean)
        If Wb Is Nothing Then
            Return
        End If

        Globals.Ribbons.TfsRibbon.SetEnabled(Wb, force, True)

        ' マルチスレッドだと不安定になるのでまた今度
        'Dim task As New System.Threading.Thread(
        '    Sub()
        '        Globals.Ribbons.TfsRibbon.SetEnabled(Wb, force, True)
        '    End Sub)

        'task.SetApartmentState(System.Threading.ApartmentState.STA)
        'task.Start()
    End Sub


    Private Function TF() As TfExe
        If Globals.Ribbons.TfsRibbon.Tf Is Nothing Then
            Globals.Ribbons.TfsRibbon.Init()
        End If
        Return Globals.Ribbons.TfsRibbon.Tf
    End Function

#End Region

End Class
