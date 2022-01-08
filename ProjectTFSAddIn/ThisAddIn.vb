﻿
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

    Private Sub Application_WindowActivate(activatedWindow As MSProject.Window) Handles Application.WindowActivate
        SetEnabled(Application.ActiveProject, False)
    End Sub

    Private Sub Application_ProjectAfterSave() Handles Application.ProjectAfterSave
        SetEnabled(Globals.ThisAddIn.ActiveDocument, True)
    End Sub

    Private Sub Application_ProjectBeforeClose(pj As MSProject.Project, ByRef Cancel As Boolean) Handles Application.ProjectBeforeClose
        TF.CloseOutputTaskPane(pj)
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
        Return Globals.ThisAddIn.Application.ActiveProject
    End Function

    Public Sub Close(Optional saveChanges As Object = Nothing) Implements IDocument.Close
        Globals.ThisAddIn.Application.ActiveWindow.Close()
    End Sub

    Public Sub CompareSideBySideWith(name As String) Implements IDocument.CompareSideBySideWith
    End Sub

    Public ReadOnly Property FullName As String Implements IDocument.FullName
        Get
            Dim filename As String = TF.GetLocalPath(Globals.ThisAddIn.Application.ActiveProject.FullName)
            Return filename
        End Get
    End Property

    Public ReadOnly Property Name As String Implements IDocument.Name
        Get
            Return Globals.ThisAddIn.Application.ActiveProject.Name
        End Get
    End Property

    Public Sub Open(filename As String) Implements IDocument.Open
        Globals.ThisAddIn.Application.FileOpenEx(filename)
    End Sub

    Public Sub Save() Implements IDocument.Save
        Globals.ThisAddIn.Application.ActiveProject.Save()
    End Sub

    Public Sub SaveAs(filename As String) Implements IDocument.SaveAs
        Globals.ThisAddIn.Application.ActiveProject.SaveAs(filename)
    End Sub

    Public Property Saved As Boolean Implements IDocument.Saved
        Get
            Return Globals.ThisAddIn.Application.ActiveProject.Saved
        End Get
        Set(value As Boolean)
        End Set
    End Property

#End Region

#Region " Method "

    Private Sub SetEnabled(Wb As MSProject.Project, ByVal force As Boolean)
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
