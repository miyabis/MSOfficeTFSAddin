
Public Class ThisAddIn
    Implements IDocument

    Private _action As IAction

    Private _open As Boolean

#Region " Handles "

#Region " ThisAddIn "

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        Try
            Globals.Ribbons.TfsRibbon.Disable()
            _action = New Action(Globals.Ribbons.TfsRibbon, Me)
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

#End Region
#Region " Application "

    Private Sub Application_WindowActivate(Wb As Microsoft.Office.Interop.Excel.Workbook, Wn As Microsoft.Office.Interop.Excel.Window) Handles Application.WindowActivate
        Try
            If _open Then
                _open = False
                Return
            End If
            SetEnabled(Wb, False)
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub Application_WorkbookOpen(Wb As Microsoft.Office.Interop.Excel.Workbook) Handles Application.WorkbookOpen
        Try
            TF.CreateOutputTaskPane(Wb)
            SetEnabled(Wb, True)
            _open = True
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub Application_WorkbookAfterSave(Wb As Microsoft.Office.Interop.Excel.Workbook, Success As Boolean) Handles Application.WorkbookAfterSave
        Try
            SetEnabled(Wb, True)
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub Application_WorkbookBeforeClose(Wb As Microsoft.Office.Interop.Excel.Workbook, ByRef Cancel As Boolean) Handles Application.WorkbookBeforeClose
        Try
            TF.CloseOutputTaskPane(Wb)
            Globals.Ribbons.TfsRibbon.RemoveEnables(Wb)
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try
    End Sub

#End Region

#End Region

#Region " Implements "

    Public Function ActiveDocument() As Object Implements IDocument.ActiveDocument
        Return Globals.ThisAddIn.Application.ActiveWorkbook
    End Function

    Public ReadOnly Property Action As IAction Implements IDocument.Action
        Get
            Return _action
        End Get
    End Property

    Public Sub Close(Optional saveChanges As Object = Nothing) Implements IDocument.Close
        Globals.ThisAddIn.Application.ActiveWorkbook.Close(saveChanges)
    End Sub

    Public ReadOnly Property FullName As String Implements IDocument.FullName
        Get
            Return Globals.ThisAddIn.Application.ActiveWorkbook.FullName
        End Get
    End Property

    Public ReadOnly Property Name As String Implements IDocument.Name
        Get
            Return Globals.ThisAddIn.Application.ActiveWorkbook.Name
        End Get
    End Property

    Public Sub Open(filename As String) Implements IDocument.Open
        Globals.ThisAddIn.Application.Workbooks.Open(filename)
    End Sub

    Public Property Saved As Boolean Implements IDocument.Saved
        Get
            Return Globals.ThisAddIn.Application.ActiveWorkbook.Saved
        End Get
        Set(value As Boolean)
            Globals.ThisAddIn.Application.ActiveWorkbook.Saved = value
        End Set
    End Property

    Public Sub Save() Implements IDocument.Save
        Globals.ThisAddIn.Application.ActiveWorkbook.Save()
    End Sub

    Public Sub SaveAs(filename As String) Implements IDocument.SaveAs
        Globals.ThisAddIn.Application.ActiveWorkbook.SaveAs(filename)
    End Sub

    Public Sub CompareSideBySideWith(name As String) Implements IDocument.CompareSideBySideWith
        Globals.ThisAddIn.Application.Windows.CompareSideBySideWith(name)
        Globals.ThisAddIn.Application.Windows.SyncScrollingSideBySide = True
        Globals.ThisAddIn.Application.Windows.Arrange(Excel.XlArrangeStyle.xlArrangeStyleVertical)
    End Sub

#End Region

#Region " Method "

    Private Sub SetEnabled(Wb As Microsoft.Office.Interop.Excel.Workbook, ByVal force As Boolean)
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

