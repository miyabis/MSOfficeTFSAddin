Partial Class TfsRibbon
	Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

	<System.Diagnostics.DebuggerNonUserCode()> _
	Public Sub New(ByVal container As System.ComponentModel.IContainer)
		MyClass.New()

		'Windows.Forms クラス作成デザイナーのサポートに必要です。
		If (container IsNot Nothing) Then
			container.Add(Me)
		End If

	End Sub

	<System.Diagnostics.DebuggerNonUserCode()> _
	Public Sub New()
		MyBase.New(Globals.Factory.GetRibbonFactory())

		'この呼び出しは、コンポーネント デザイナーで必要です。
		InitializeComponent()

	End Sub

	'Component は、コンポーネント一覧に後処理を実行するために dispose をオーバーライドします。
	<System.Diagnostics.DebuggerNonUserCode()> _
	Protected Overrides Sub Dispose(ByVal disposing As Boolean)
		Try
			If disposing AndAlso components IsNot Nothing Then
				components.Dispose()
			End If
		Finally
			MyBase.Dispose(disposing)
		End Try
	End Sub

	'コンポーネント デザイナーで必要です。
	Private components As System.ComponentModel.IContainer

	'メモ: 以下のプロシージャはコンポーネント デザイナーで必要です。
	'これはコンポーネント デザイナーを使用して変更できます。
	'コード エディターを使用して変更しないでください。
	<System.Diagnostics.DebuggerStepThrough()> _
	Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(TfsRibbon))
        Me.tabTfs = Me.Factory.CreateRibbonTab
        Me.groupExt = Me.Factory.CreateRibbonGroup
        Me.btnVS = Me.Factory.CreateRibbonButton
        Me.btnExplorer = Me.Factory.CreateRibbonButton
        Me.btnWebAccess = Me.Factory.CreateRibbonButton
        Me.groupAction = Me.Factory.CreateRibbonGroup
        Me.btnGetItem = Me.Factory.CreateRibbonButton
        Me.btnGetItemFolder = Me.Factory.CreateRibbonButton
        Me.btnGetWorkspace = Me.Factory.CreateRibbonButton
        Me.Separator3 = Me.Factory.CreateRibbonSeparator
        Me.btnCheckOut = Me.Factory.CreateRibbonButton
        Me.btnAdd = Me.Factory.CreateRibbonButton
        Me.btnRename = Me.Factory.CreateRibbonButton
        Me.btnUndo = Me.Factory.CreateRibbonButton
        Me.Separator2 = Me.Factory.CreateRibbonSeparator
        Me.btnCheckIn = Me.Factory.CreateRibbonButton
        Me.btnShelve = Me.Factory.CreateRibbonButton
        Me.btnUnshelve = Me.Factory.CreateRibbonButton
        Me.Separator4 = Me.Factory.CreateRibbonSeparator
        Me.btnResolveByCopy = Me.Factory.CreateRibbonButton
        Me.btnResolve = Me.Factory.CreateRibbonButton
        Me.Separator1 = Me.Factory.CreateRibbonSeparator
        Me.btnHistory = Me.Factory.CreateRibbonButton
        Me.btnDifference = Me.Factory.CreateRibbonButton
        Me.btnInfo = Me.Factory.CreateRibbonButton
        Me.groupWindow = Me.Factory.CreateRibbonGroup
        Me.btnOutputPane = Me.Factory.CreateRibbonButton
        Me.groupVba = Me.Factory.CreateRibbonGroup
        Me.btnScriptExport = Me.Factory.CreateRibbonButton
        Me.tabTfs.SuspendLayout()
        Me.groupExt.SuspendLayout()
        Me.groupAction.SuspendLayout()
        Me.groupWindow.SuspendLayout()
        Me.groupVba.SuspendLayout()
        Me.SuspendLayout()
        '
        'tabTfs
        '
        Me.tabTfs.Groups.Add(Me.groupExt)
        Me.tabTfs.Groups.Add(Me.groupVba)
        Me.tabTfs.Groups.Add(Me.groupAction)
        Me.tabTfs.Groups.Add(Me.groupWindow)
        resources.ApplyResources(Me.tabTfs, "tabTfs")
        Me.tabTfs.Name = "tabTfs"
        '
        'groupExt
        '
        Me.groupExt.Items.Add(Me.btnVS)
        Me.groupExt.Items.Add(Me.btnExplorer)
        Me.groupExt.Items.Add(Me.btnWebAccess)
        resources.ApplyResources(Me.groupExt, "groupExt")
        Me.groupExt.Name = "groupExt"
        '
        'btnVS
        '
        Me.btnVS.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        resources.ApplyResources(Me.btnVS, "btnVS")
        Me.btnVS.Name = "btnVS"
        Me.btnVS.ShowImage = True
        '
        'btnExplorer
        '
        Me.btnExplorer.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        resources.ApplyResources(Me.btnExplorer, "btnExplorer")
        Me.btnExplorer.Name = "btnExplorer"
        Me.btnExplorer.ShowImage = True
        '
        'btnWebAccess
        '
        Me.btnWebAccess.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        resources.ApplyResources(Me.btnWebAccess, "btnWebAccess")
        Me.btnWebAccess.Name = "btnWebAccess"
        Me.btnWebAccess.ShowImage = True
        '
        'groupAction
        '
        Me.groupAction.Items.Add(Me.btnGetItem)
        Me.groupAction.Items.Add(Me.btnGetItemFolder)
        Me.groupAction.Items.Add(Me.btnGetWorkspace)
        Me.groupAction.Items.Add(Me.Separator3)
        Me.groupAction.Items.Add(Me.btnCheckOut)
        Me.groupAction.Items.Add(Me.btnAdd)
        Me.groupAction.Items.Add(Me.btnRename)
        Me.groupAction.Items.Add(Me.btnUndo)
        Me.groupAction.Items.Add(Me.Separator2)
        Me.groupAction.Items.Add(Me.btnCheckIn)
        Me.groupAction.Items.Add(Me.btnShelve)
        Me.groupAction.Items.Add(Me.btnUnshelve)
        Me.groupAction.Items.Add(Me.Separator4)
        Me.groupAction.Items.Add(Me.btnResolveByCopy)
        Me.groupAction.Items.Add(Me.btnResolve)
        Me.groupAction.Items.Add(Me.Separator1)
        Me.groupAction.Items.Add(Me.btnHistory)
        Me.groupAction.Items.Add(Me.btnDifference)
        Me.groupAction.Items.Add(Me.btnInfo)
        resources.ApplyResources(Me.groupAction, "groupAction")
        Me.groupAction.Name = "groupAction"
        '
        'btnGetItem
        '
        resources.ApplyResources(Me.btnGetItem, "btnGetItem")
        Me.btnGetItem.Name = "btnGetItem"
        Me.btnGetItem.ShowImage = True
        '
        'btnGetItemFolder
        '
        resources.ApplyResources(Me.btnGetItemFolder, "btnGetItemFolder")
        Me.btnGetItemFolder.Name = "btnGetItemFolder"
        Me.btnGetItemFolder.ShowImage = True
        '
        'btnGetWorkspace
        '
        resources.ApplyResources(Me.btnGetWorkspace, "btnGetWorkspace")
        Me.btnGetWorkspace.Name = "btnGetWorkspace"
        Me.btnGetWorkspace.ShowImage = True
        '
        'Separator3
        '
        Me.Separator3.Name = "Separator3"
        '
        'btnCheckOut
        '
        Me.btnCheckOut.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        resources.ApplyResources(Me.btnCheckOut, "btnCheckOut")
        Me.btnCheckOut.Name = "btnCheckOut"
        Me.btnCheckOut.ShowImage = True
        '
        'btnAdd
        '
        resources.ApplyResources(Me.btnAdd, "btnAdd")
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.ShowImage = True
        '
        'btnRename
        '
        resources.ApplyResources(Me.btnRename, "btnRename")
        Me.btnRename.Name = "btnRename"
        Me.btnRename.ShowImage = True
        '
        'btnUndo
        '
        resources.ApplyResources(Me.btnUndo, "btnUndo")
        Me.btnUndo.Name = "btnUndo"
        Me.btnUndo.ShowImage = True
        '
        'Separator2
        '
        Me.Separator2.Name = "Separator2"
        '
        'btnCheckIn
        '
        Me.btnCheckIn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        resources.ApplyResources(Me.btnCheckIn, "btnCheckIn")
        Me.btnCheckIn.Name = "btnCheckIn"
        Me.btnCheckIn.ShowImage = True
        '
        'btnShelve
        '
        resources.ApplyResources(Me.btnShelve, "btnShelve")
        Me.btnShelve.Name = "btnShelve"
        Me.btnShelve.ShowImage = True
        '
        'btnUnshelve
        '
        resources.ApplyResources(Me.btnUnshelve, "btnUnshelve")
        Me.btnUnshelve.Name = "btnUnshelve"
        Me.btnUnshelve.ShowImage = True
        '
        'Separator4
        '
        Me.Separator4.Name = "Separator4"
        '
        'btnResolveByCopy
        '
        resources.ApplyResources(Me.btnResolveByCopy, "btnResolveByCopy")
        Me.btnResolveByCopy.Name = "btnResolveByCopy"
        Me.btnResolveByCopy.ShowImage = True
        '
        'btnResolve
        '
        resources.ApplyResources(Me.btnResolve, "btnResolve")
        Me.btnResolve.Name = "btnResolve"
        Me.btnResolve.ShowImage = True
        '
        'Separator1
        '
        Me.Separator1.Name = "Separator1"
        '
        'btnHistory
        '
        resources.ApplyResources(Me.btnHistory, "btnHistory")
        Me.btnHistory.Name = "btnHistory"
        Me.btnHistory.ShowImage = True
        '
        'btnDifference
        '
        resources.ApplyResources(Me.btnDifference, "btnDifference")
        Me.btnDifference.Name = "btnDifference"
        Me.btnDifference.ShowImage = True
        '
        'btnInfo
        '
        resources.ApplyResources(Me.btnInfo, "btnInfo")
        Me.btnInfo.Name = "btnInfo"
        Me.btnInfo.ShowImage = True
        '
        'groupWindow
        '
        Me.groupWindow.Items.Add(Me.btnOutputPane)
        resources.ApplyResources(Me.groupWindow, "groupWindow")
        Me.groupWindow.Name = "groupWindow"
        '
        'btnOutputPane
        '
        resources.ApplyResources(Me.btnOutputPane, "btnOutputPane")
        Me.btnOutputPane.Name = "btnOutputPane"
        Me.btnOutputPane.ShowImage = True
        '
        'groupVba
        '
        Me.groupVba.Items.Add(Me.btnScriptExport)
        resources.ApplyResources(Me.groupVba, "groupVba")
        Me.groupVba.Name = "groupVba"
        '
        'btnScriptExport
        '
        resources.ApplyResources(Me.btnScriptExport, "btnScriptExport")
        Me.btnScriptExport.Name = "btnScriptExport"
        Me.btnScriptExport.ShowImage = True
        '
        'TfsRibbon
        '
        Me.Name = "TfsRibbon"
        Me.RibbonType = "Microsoft.PowerPoint.Presentation"
        Me.Tabs.Add(Me.tabTfs)
        Me.tabTfs.ResumeLayout(False)
        Me.tabTfs.PerformLayout()
        Me.groupExt.ResumeLayout(False)
        Me.groupExt.PerformLayout()
        Me.groupAction.ResumeLayout(False)
        Me.groupAction.PerformLayout()
        Me.groupWindow.ResumeLayout(False)
        Me.groupWindow.PerformLayout()
        Me.groupVba.ResumeLayout(False)
        Me.groupVba.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents tabTfs As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents groupAction As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnCheckIn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnCheckOut As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnAdd As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnUndo As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnInfo As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnDifference As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnResolve As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents groupWindow As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Separator1 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents btnExplorer As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnVS As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnGetItem As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnGetWorkspace As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator3 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents btnHistory As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnShelve As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnUnshelve As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator2 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents btnRename As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnResolveByCopy As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator4 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents groupExt As Microsoft.Office.Tools.Ribbon.RibbonGroup
	Friend WithEvents btnWebAccess As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnGetItemFolder As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnOutputPane As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents groupVba As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnScriptExport As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property TfsRibbon() As TfsRibbon
        Get
            Return Me.GetRibbon(Of TfsRibbon)()
        End Get
    End Property
End Class
