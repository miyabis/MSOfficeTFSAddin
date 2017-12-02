
Public Delegate Sub ExecuteMethod(ByVal context As ActionContext)

Public Interface IAction

    ''' <summary>
    ''' 実行
    ''' </summary>
    ''' <param name="method"></param>
    ''' <remarks></remarks>
    Sub Execute(ByVal method As ExecuteMethod)

    ''' <summary>
    ''' 実行（閉じて開きなおす）
    ''' </summary>
    ''' <param name="method"></param>
    ''' <remarks></remarks>
    Sub ExecuteCloseOpen(ByVal method As ExecuteMethod)

End Interface
