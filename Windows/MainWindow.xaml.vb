Imports System.ComponentModel
Imports System.Runtime.CompilerServices
Imports System.Windows.Threading

Class MainWindow

    ''' <summary>
    ''' AddinBaseClass object to be used referenced by this window.
    ''' </summary>
    Private _Addin As AddinBaseClass

    ''' <summary>
    ''' Just a sample event to show how to hook back into the MainWindowControlDef. <para/>
    ''' This can be modified to fit your needs, or deleted.
    ''' </summary>
    Public Event MyEvent()

    ''' <summary>
    ''' Creates a new instance of the window and moves it to the last saved location.
    ''' </summary>
    ''' <param name="Addin">[AddinBaseClass] AddinBaseClass object to be used referenced by this window.</param>
    Sub New(ByVal Addin As AddinBaseClass)

        ' This call is required by the designer.
        InitializeComponent()

        'Link the Add-in Object
        _Addin = Addin

        'Set Inventor as the parent of this window.
        'This will allow your window to be minimized with Inventor.
        Call _Addin.SetInventorAsOwnerWindow(Me)

        'Set the saved window location.
        Me.Top = My.Settings.WindowLocation.X
        Me.Left = My.Settings.WindowLocation.Y


    End Sub

    ''' <summary>
    ''' Closes the window and saves its last location on screen to the app settings.
    ''' </summary>
    Private Sub OnCloseWindow(target As Object, e As ExecutedRoutedEventArgs)
        My.Settings.WindowLocation = New System.Drawing.Point(Me.Top, Me.Left)
        My.Settings.Save()
        SystemCommands.CloseWindow(Me)
    End Sub

    ''' <summary>
    ''' Empty delegate so we can force a window update o the UI thread from other threads.
    ''' </summary>
    Public EmptyDelegate As Action = Sub()
                                     End Sub

    Public Sub Refresh()
        Dispatcher.Invoke(DispatcherPriority.Render, EmptyDelegate)
    End Sub

End Class