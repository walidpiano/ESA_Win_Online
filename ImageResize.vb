Imports System
Imports System.Collections.Generic
Imports System.Web
Imports System.Drawing
Imports System.IO
Imports System.Drawing.Drawing2D
Imports System.Drawing.Imaging

Namespace SharedKernel.Core.Utilities

    Module ImageUtilities

        Private Function GetEncoder(ByVal format As ImageFormat) As ImageCodecInfo
            Dim codecs As ImageCodecInfo() = ImageCodecInfo.GetImageDecoders()
            For Each codec As ImageCodecInfo In codecs
                If codec.FormatID = format.Guid Then
                    Return codec
                End If
            Next

            Return Nothing
        End Function

        Private Sub saveImageToLocation(ByVal theImage As Image, ByVal saveLocation As String)
            Dim saveFolder As String = Path.GetDirectoryName(saveLocation)
            If Not Directory.Exists(saveFolder) Then
                Directory.CreateDirectory(saveFolder)
            End If

            Dim extension As String = Path.GetExtension(saveLocation)
            Dim imageEncoder As ImageCodecInfo
            Select Case extension.ToLower()
                Case ".png"
                    imageEncoder = GetEncoder(ImageFormat.Png)
                Case ".gif"
                    imageEncoder = GetEncoder(ImageFormat.Gif)
                Case Else
                    imageEncoder = GetEncoder(ImageFormat.Jpeg)
            End Select

            Dim myEncoder As System.Drawing.Imaging.Encoder = System.Drawing.Imaging.Encoder.Quality
            Dim myEncoderParameters As EncoderParameters = New EncoderParameters(1)
            Dim myEncoderParameter As EncoderParameter = New EncoderParameter(myEncoder, 85L)
            myEncoderParameters.Param(0) = myEncoderParameter
            theImage.Save(saveLocation, imageEncoder, myEncoderParameters)
        End Sub

        Sub resizeImageAndSave(ByVal ImageToResize As Image, ByVal newWidth As Integer, ByVal maxHeight As Integer, ByVal onlyResizeIfWider As Boolean, ByVal thumbnailSaveAs As String)
            Dim thumbnail As Image = resizeImage(ImageToResize, newWidth, maxHeight, onlyResizeIfWider)
            saveImageToLocation(thumbnail, thumbnailSaveAs)
        End Sub

        Sub resizeImageAndSave(ByVal ImageToResize As Image, ByVal newWidth As Integer, ByVal maxHeight As Integer, ByVal onlyResizeIfWider As Boolean, ByVal thumbnailSaveAs As String, ByVal directory As String)
            Dim thumbnail As Image = resizeImage(ImageToResize, newWidth, maxHeight, onlyResizeIfWider)
            If IO.Directory.Exists(directory) = False Then
                IO.Directory.CreateDirectory(directory)
            End If

            saveImageToLocation(thumbnail, thumbnailSaveAs)
        End Sub

        Sub resizeImageAndSave(ByVal imageLocation As String, ByVal newWidth As Integer, ByVal maxHeight As Integer, ByVal onlyResizeIfWider As Boolean, ByVal thumbnailSaveAs As String)
            Dim loadedImage As Image = Image.FromFile(imageLocation)
            Dim thumbnail As Image = resizeImage(loadedImage, newWidth, maxHeight, onlyResizeIfWider)
            saveImageToLocation(thumbnail, thumbnailSaveAs)
        End Sub

        Sub resizeImageAndSave(ByVal imageContents As Byte(), ByVal newWidth As Integer, ByVal maxHeight As Integer, ByVal onlyResizeIfWider As Boolean, ByVal thumbnailSaveAs As String)
            Dim loadedImage As Image
            Using ms As MemoryStream = New MemoryStream(imageContents)
                loadedImage = System.Drawing.Image.FromStream(ms)
            End Using

            Dim thumbnail As Image = resizeImage(loadedImage, newWidth, maxHeight, onlyResizeIfWider)
            saveImageToLocation(thumbnail, thumbnailSaveAs)
        End Sub

        Function CropImage(ByVal imageBytes As Byte(), ByVal topLeftX As Integer, ByVal topLeftY As Integer, ByVal height As Integer, ByVal width As Integer) As Byte()
            Dim src, target As Bitmap
            Dim jgpEncoder As ImageCodecInfo = GetEncoder(ImageFormat.Jpeg)
            Dim myEncoder As System.Drawing.Imaging.Encoder = System.Drawing.Imaging.Encoder.Quality
            Dim myEncoderParameters As EncoderParameters = New EncoderParameters(1)
            Dim myEncoderParameter As EncoderParameter = New EncoderParameter(myEncoder, 80L)
            myEncoderParameters.Param(0) = myEncoderParameter
            Dim cropRect As Rectangle = New Rectangle(topLeftX, topLeftY, width, height)
            Using ms As MemoryStream = New MemoryStream(imageBytes)
                src = TryCast(System.Drawing.Image.FromStream(ms), Bitmap)
            End Using

            target = New Bitmap(cropRect.Width, cropRect.Height, PixelFormat.Format24bppRgb)
            target.SetResolution(72, 72)
            Using graphics As Graphics = Graphics.FromImage(target)
                graphics.SmoothingMode = SmoothingMode.AntiAlias
                graphics.InterpolationMode = InterpolationMode.HighQualityBicubic
                graphics.PixelOffsetMode = PixelOffsetMode.HighQuality
                graphics.FillRectangle(New System.Drawing.SolidBrush(System.Drawing.Color.White), New System.Drawing.Rectangle(0, 0, target.Width, target.Height))
                graphics.DrawImage(src, New Rectangle(0, 0, target.Width, target.Height), cropRect, GraphicsUnit.Pixel)
            End Using

            Dim memoryStream As MemoryStream = New MemoryStream()
            target.Save(memoryStream, jgpEncoder, myEncoderParameters)
            Return memoryStream.GetBuffer()
        End Function

        Function resizeImage(ByVal ImageToResize As Image, ByVal newWidth As Integer, ByVal maxHeight As Integer, ByVal onlyResizeIfWider As Boolean) As Image
            Try
                ImageToResize.RotateFlip(System.Drawing.RotateFlipType.Rotate180FlipNone)
                ImageToResize.RotateFlip(System.Drawing.RotateFlipType.Rotate180FlipNone)
                If onlyResizeIfWider Then
                    If ImageToResize.Width <= newWidth Then
                        newWidth = ImageToResize.Width
                    End If
                End If

                Dim newHeight As Integer = ImageToResize.Height * newWidth / ImageToResize.Width
                If newHeight > maxHeight Then
                    newWidth = ImageToResize.Width * maxHeight / ImageToResize.Height
                    newHeight = maxHeight
                End If

                Dim newImage As Bitmap = New Bitmap(newWidth, newHeight, PixelFormat.Format32bppArgb)
                newImage.SetResolution(72, 72)
                newImage.SetResolution(ImageToResize.HorizontalResolution, ImageToResize.VerticalResolution)
                Using gr As Graphics = Graphics.FromImage(newImage)
                    gr.SmoothingMode = SmoothingMode.HighQuality
                    gr.InterpolationMode = InterpolationMode.HighQualityBicubic
                    gr.CompositingMode = CompositingMode.SourceCopy
                    gr.PixelOffsetMode = PixelOffsetMode.HighQuality
                    gr.DrawImage(ImageToResize, New Rectangle(0, 0, newWidth, newHeight))
                End Using

                Return newImage
            Catch
                Return Nothing
            Finally
                ImageToResize.Dispose()
            End Try
        End Function

        Function resizeImage(ByVal imageLocation As String, ByVal newWidth As Integer, ByVal maxHeight As Integer, ByVal onlyResizeIfWider As Boolean) As Image
            Dim loadedImage As Image = Image.FromFile(imageLocation)
            Return resizeImage(loadedImage, newWidth, maxHeight, onlyResizeIfWider)
        End Function

        Function resizeImageRespectHeight(ByVal ImageToResize As Image, ByVal newHeight As Integer) As Image
            Try
                ImageToResize.RotateFlip(System.Drawing.RotateFlipType.Rotate180FlipNone)
                ImageToResize.RotateFlip(System.Drawing.RotateFlipType.Rotate180FlipNone)
                Dim newWidth As Integer = ImageToResize.Width * newHeight / ImageToResize.Height
                Dim newImage As Bitmap = New Bitmap(newWidth, newHeight, System.Drawing.Imaging.PixelFormat.Format24bppRgb)
                newImage.SetResolution(ImageToResize.HorizontalResolution, ImageToResize.VerticalResolution)
                Using gr As Graphics = Graphics.FromImage(newImage)
                    gr.SmoothingMode = SmoothingMode.HighQuality
                    gr.InterpolationMode = InterpolationMode.HighQualityBicubic
                    gr.CompositingMode = CompositingMode.SourceCopy
                    gr.PixelOffsetMode = PixelOffsetMode.HighQuality
                    gr.DrawImage(ImageToResize, New Rectangle(0, 0, newWidth, newHeight))
                End Using

                Return newImage
            Catch
                Return Nothing
            Finally
                ImageToResize.Dispose()
            End Try
        End Function
    End Module
End Namespace
