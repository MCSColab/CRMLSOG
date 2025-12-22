Imports System.Media

Class Application
    Private fxCommon As New Common()
    Protected Overrides Sub OnStartup(ByVal e As StartupEventArgs)
        MyBase.OnStartup(e)

        ' Attach the event handler for unhandled exceptions
        AddHandler AppDomain.CurrentDomain.UnhandledException, AddressOf CurrentDomain_UnhandledException
    End Sub
    ' Application-level events, such as Startup, Exit, and DispatcherUnhandledException
    ' can be handled in this file.
    Private Sub CurrentDomain_UnhandledException(ByVal sender As Object, ByVal e As UnhandledExceptionEventArgs)
        Dim exception As Exception = CType(e.ExceptionObject, Exception)
        fxCommon.GenerateLog(exception)
        ' Handle the exception here, 'exception' contains the exception object
        SystemSounds.Exclamation.Play()
        MessageBox.Show($"An unhandled exception occurred: {exception.Message}", "Unhandled Exception")


    End Sub
End Class
