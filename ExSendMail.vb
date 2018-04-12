Imports System.Net.Mail
Imports System.IO
Imports System.Drawing.Imaging

Public Class ExSendMail
    Public Shared Function SendToESA() As String
        Dim result As String

        Dim Status As String
        Dim Photo As Image = Nothing
        Dim DataLine As String = ""

        If frmMain.tsStatus.IsOn = False Then
            Status = "NEW"
            If Not frmMain.pePhotoN.Image Is Nothing Then
                Photo = ExClass.LowerImage(frmMain.pePhotoN.Image)
            End If
            DataLine = ExClass.ConvertToLine(True, False)
        Else
            Status = "Old"
            If Not frmMain.pePhotoO.Image Is Nothing Then
                Photo = ExClass.LowerImage(frmMain.pePhotoO.Image)
            End If
            DataLine = ExClass.ConvertToLine(False, False)
        End If
        Dim stream As New MemoryStream()
        If Not Photo Is Nothing Then
            'Photo.Save("Photo.jpg")
            Photo.Save(stream, ImageFormat.Jpeg)
            stream.Position = 0
            'ElseIf IO.File.Exists("Photo.jpg") Then
            '    Try
            'IO.File.Delete("Photo.jpg")
            'Catch ex As Exception

            'End Try
        End If
        

        ''' var stream = new MemorySteam();
        ''' image.save(stream, ImageFormat.Jpeg);
        ''' stream.Position = 0;
        ''' mail.Attachment.Add(new Attachment(stream, "image/jpg"));

        Dim theMail As New MailMessage()
        Dim theSmtpServer As New SmtpClient("smtp.mail.yahoo.com")
        theMail.From = New MailAddress("esaegypt@yahoo.com")
        theMail.[To].Add("esaegypt@yahoo.com")
        theMail.Subject = "ESA Registration (" & Status & ")"
        theMail.Body = DataLine

        If Not Photo Is Nothing Then
            'Dim attch As System.Net.Mail.Attachment
            'attch = New System.Net.Mail.Attachment("Photo.jpg")
            'theMail.Attachments.Add(attch)
            theMail.Attachments.Add(New Attachment(stream, "Photo.jpg"))
        End If


        theSmtpServer.Port = 587
        theSmtpServer.Credentials = New System.Net.NetworkCredential("esaegypt@yahoo.com", "eessaa2018")
        theSmtpServer.EnableSsl = True

        Try
            theSmtpServer.Send(theMail)
            result = "Registration has been successfully sent!"
        Catch ex As Exception
            result = "Failed to send your registraion. Please check your internet connection, and try again!"
        End Try

        Return result

    End Function
End Class
