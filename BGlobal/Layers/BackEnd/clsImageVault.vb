Imports System.IO
Imports System.Windows.Media.Imaging

Public Class clsImageVault
    Implements IDisposable

#Region "Declarations"
    Private objImages As List(Of clsImageVaultItem)
#End Region

#Region "Constructor"
    Public Sub New()
        Try
            objImages = New List(Of clsImageVaultItem)
        Catch ex As Exception
            Throw
        End Try
    End Sub
#End Region

#Region "Functions"
    Public Function OpenImage(p_strfPath As String) As BitmapImage

        Dim objAuxItem As clsImageVaultItem

        Try

            objAuxItem = objImages.Find(Function(img As clsImageVaultItem) img.ImagePath.Trim.ToUpper = p_strfPath.Trim.ToUpper)
            If objAuxItem IsNot Nothing Then
                objImages.Remove(objAuxItem)
            End If

            objAuxItem = New clsImageVaultItem(p_strfPath)

            objImages.Add(objAuxItem)

            Return objAuxItem.Image

        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Sub RemoveImage(p_strfPath As String)

        Dim objAuxItem As clsImageVaultItem

        Try

            objAuxItem = objImages.Find(Function(img As clsImageVaultItem) img.ImagePath.Trim.ToUpper = p_strfPath.Trim.ToUpper)
            If objAuxItem IsNot Nothing Then
                objImages.Remove(objAuxItem)
            End If

        Catch ex As Exception
            Throw
        End Try
    End Sub

    Public Function GetImageByPath(p_strfPath As String) As BitmapImage

        Dim objReturn As BitmapImage
        Dim objAuxItem As clsImageVaultItem

        Try

            objReturn = Nothing

            objAuxItem = objImages.Find(Function(img As clsImageVaultItem) img.ImagePath.Trim.ToUpper = p_strfPath.Trim.ToUpper)
            If objAuxItem IsNot Nothing Then
                objReturn = objAuxItem.Image
            Else
                Throw New FileNotFoundException("Image not found.", p_strfPath)
            End If

            Return objReturn

        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function GetImageByName(p_strfName As String) As BitmapImage

        Dim objReturn As BitmapImage
        Dim objAuxItem As clsImageVaultItem

        Try

            objReturn = Nothing

            objAuxItem = objImages.Find(Function(img As clsImageVaultItem) img.ImageName.Trim.ToUpper = p_strfName.Trim.ToUpper)
            If objAuxItem IsNot Nothing Then
                objReturn = objAuxItem.Image
            Else
                Throw New FileNotFoundException("Image not found.", p_strfName)
            End If

            Return objReturn

        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Sub Clear()
        Try
            Call Dispose()
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        Try
            For Each Image As clsImageVaultItem In objImages
                Image.Dispose()
            Next
            objImages.Clear()
        Catch ex As Exception
            Throw
        End Try
    End Sub
#End Region

#Region "Properties"
    Public Property Images() As List(Of clsImageVaultItem)
        Get
            Return objImages
        End Get
        Set(ByVal value As List(Of clsImageVaultItem))
            objImages = value
        End Set
    End Property
#End Region
End Class

Public Class clsImageVaultItem
    Implements IDisposable

#Region "Declarations"
    Private strImageName As String
    Private strImagePath As String
    Private objStream As MemoryStream
    Private objImage As BitmapImage
#End Region

#Region "Constructor"
    Public Sub New(p_strPath As String)
        Try
            Call sbLoadImage(p_strPath)
        Catch ex As Exception
            Throw
        End Try
    End Sub
#End Region

#Region "Functions"
    Private Sub sbLoadImage(p_strPath As String)

        Dim objAuxImage As BitmapImage
        Dim objEncoder As BitmapEncoder

        Try

            strImagePath = p_strPath
            strImageName = Path.GetFileName(p_strPath)

            objAuxImage = New BitmapImage()
            objAuxImage.BeginInit()
            objAuxImage.UriSource = New Uri(Path.GetFullPath(strImagePath), UriKind.Absolute)
            objAuxImage.EndInit()


            Try
                objEncoder = New PngBitmapEncoder
                objEncoder.Frames.Add(BitmapFrame.Create(objAuxImage))
                objStream = New MemoryStream
                objEncoder.Save(objStream)
            Catch ex As Exception
                Try
                    objEncoder = New JpegBitmapEncoder
                    objEncoder.Frames.Add(BitmapFrame.Create(objAuxImage))
                    objStream = New MemoryStream
                    objEncoder.Save(objStream)
                Catch ex2 As Exception
                    objEncoder = New BmpBitmapEncoder
                    objEncoder.Frames.Add(BitmapFrame.Create(objAuxImage))
                    objStream = New MemoryStream
                    objEncoder.Save(objStream)
                End Try
            End Try

            objEncoder = Nothing
            objAuxImage = Nothing

            objImage = New BitmapImage
            objImage.BeginInit()
            objImage.StreamSource = objStream
            objImage.CacheOption = BitmapCacheOption.OnLoad
            objImage.EndInit()
            objImage.Freeze()

        Catch ex As Exception
            Throw
        End Try
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        Try
            objStream.Close()
            objStream.Dispose()
            objStream = Nothing
            objImage = Nothing
            strImagePath = ""
            strImageName = ""
        Catch ex As Exception
            Throw
        End Try
    End Sub
#End Region

#Region "Properties"
    Public ReadOnly Property ImagePath() As String
        Get
            Return strImagePath
        End Get
    End Property

    Public ReadOnly Property Stream As MemoryStream
        Get
            Return objStream
        End Get
    End Property

    Public ReadOnly Property Image() As BitmapImage
        Get
            Return objImage
        End Get
    End Property

    Public ReadOnly Property ImageName As String
        Get
            Return strImageName
        End Get
    End Property
#End Region
End Class
