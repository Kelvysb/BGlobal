Imports BGlobal

Class MainWindow

    Private Sub MainWindow_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        Try
            BLogger.sbInitialize(System.AppDomain.CurrentDomain.BaseDirectory & "Log\")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub btnTestDirect_Click(sender As Object, e As RoutedEventArgs) Handles btnTestDirect.Click
        Try
            BLogger.Instance.sbLog("Test", "1", "Direct Error test", "MainWindow", "btnTestDirect_Click", 15, True)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub btnTestResume_Click(sender As Object, e As RoutedEventArgs) Handles btnTestResume.Click
        Try
            BLogger.Instance.sbLog("Test", "1", "Direct Error test", True)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub btnTestException_Click(sender As Object, e As RoutedEventArgs) Handles btnTestException.Click
        Try

            BLogger.Instance.sbLog(New Exception("Main text", New Exception("Inner code")))

            Dim intX As Integer = 0
            Dim intY As Integer = 100

            intX = intY / intX

        Catch ex As Exception
            BLogger.Instance.sbLog(ex)
        End Try
    End Sub
End Class
