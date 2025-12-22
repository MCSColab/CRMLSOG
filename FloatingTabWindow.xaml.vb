Imports System.Collections.ObjectModel

Public Class FloatingTabWindow


    Private Sub Window_Closed(sender As Object, e As EventArgs)

        'FloatTabItem(FloatingTabControl.Items(0))
    End Sub

    Private Sub Window_Closing(sender As Object, e As ComponentModel.CancelEventArgs)
        For Each win As Window In Application.Current.Windows
            If Not String.IsNullOrEmpty(win.Title) Then
                If win.Title = "MainWindow" Then
                    Dim mainWindow As MainWindow = Application.Current.Windows.OfType(Of MainWindow)().FirstOrDefault()
                    Dim tab As System.Windows.Controls.TabItem = FloatingTabControl.Items(0)
                    mainWindow.DragButton.Visibility = Visibility.Visible
                    FloatingTabControl.Items.Remove(tab)
                    mainWindow.TabControlMain.Items.Insert(mainWindow.DragTabIndex, tab)
                    mainWindow.TabControlMain.SelectedIndex = mainWindow.DragTabIndex
                End If
            End If
        Next
    End Sub
End Class
