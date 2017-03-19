''' <summary>
''' A command whose sole purpose is to
''' relay its functionality to other
''' objects by invoking delegates. The
''' default return value for the CanExecute
''' method is 'true'.
'''
''' Code converted from the example on the following link.
''' http://stackoverflow.com/questions/3531772/binding-button-click-to-a-method
''' </summary>
Public Class RelayCommand : Implements ICommand
#Region "Fields"
    ReadOnly _execute As Action(Of Object)
    ReadOnly _canExecute As Predicate(Of Object)
    Private Event ICommand_CanExecuteChanged As EventHandler Implements ICommand.CanExecuteChanged
#End Region

#Region "Constructors"
    ''' <summary>
    ''' Creates a new command that can always execute.
    ''' </summary>
    ''' <param name="execute">The execution logic.</param>
    Public Sub New(execute As Action(Of Object))
        Me.New(execute, Nothing)
    End Sub

    ''' <summary>
    ''' Creates a new command.
    ''' </summary>
    ''' <param name="execute">The execution logic.</param>
    ''' <param name="canExecute">The execution status logic.</param>
    Public Sub New(execute As Action(Of Object), canExecute As Predicate(Of Object))
        If execute Is Nothing Then
            Throw New ArgumentNullException("execute")
        End If
        _execute = execute
        _canExecute = canExecute
    End Sub
#End Region

#Region "ICommand Members"
    <DebuggerStepThrough>
    Public Function CanExecute(parameters As Object) As Boolean Implements ICommand.CanExecute
        Return If(_canExecute Is Nothing, True, _canExecute(parameters))
    End Function

    Public Custom Event CanExecuteChanged As EventHandler
        AddHandler(ByVal value As EventHandler)
            AddHandler CommandManager.RequerySuggested, value
        End AddHandler
        RemoveHandler(ByVal value As EventHandler)
            RemoveHandler CommandManager.RequerySuggested, value
        End RemoveHandler
        RaiseEvent(ByVal sender As Object, ByVal e As EventArgs)
            If (_canExecute IsNot Nothing) Then
                _canExecute.Invoke(sender)
            End If
        End RaiseEvent
    End Event

    Public Sub Execute(parameters As Object) Implements ICommand.Execute
        _execute(parameters)
    End Sub
#End Region
End Class