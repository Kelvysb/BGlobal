Imports BGlobal
Imports Microsoft.WindowsAPICodePack.Dialogs
Imports System.IO

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

    Private Sub btnImage_Click(sender As Object, e As RoutedEventArgs) Handles btnImage.Click
        Try
            Call sbOpenImage()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub sbOpenImage()
        Dim objDialog As CommonOpenFileDialog
        Dim objVault As clsImageVault

        Try
            objDialog = New CommonOpenFileDialog
            objDialog.Title = "Open"
            objDialog.Filters.Add(New CommonFileDialogFilter("Iamges", ".png;.jpg;.gif;.bmp;.jpeg"))

            If objDialog.ShowDialog = CommonFileDialogResult.Ok Then
                If Path.GetExtension(objDialog.FileName).Trim.ToUpper = ".PNG" Or
                Path.GetExtension(objDialog.FileName).Trim.ToUpper = ".JPG" Or
                Path.GetExtension(objDialog.FileName).Trim.ToUpper = ".GIF" Or
                Path.GetExtension(objDialog.FileName).Trim.ToUpper = ".BMP" Or
                Path.GetExtension(objDialog.FileName).Trim.ToUpper = ".JPEG" Then

                    objVault = New clsImageVault
                    imgTest.Source = objVault.OpenImage(objDialog.FileName)

                Else
                    MessageBox.Show("Invalid type")
                End If
            End If

            objDialog = Nothing

        Catch ex As Exception
            Throw
        End Try
    End Sub
End Class
