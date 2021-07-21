Imports System.Net.Mail

Module modEmail

    ' Send Mail Function copied from: http://vb.net-informations.com/communications/vb.net_smtp_mail.htm
    Public Function SendMail(strTO As String, strFrom As String, strSubject As String, strBody As String, strUsername As String, strPassword As String, strAttachmentPath As String)
        Try
            Dim SmtpServer As New SmtpClient()
            Dim mail As New MailMessage()

            ' Add Email ID and Password of gmail account.
            ' This will be used to send the email from
            ' In this GMail account one need to  turn on from setting -> Allow less Secure App

            SmtpServer.Port = 587
            SmtpServer.Host = "smtp.gmail.com"
            ' Citation start: https://stackoverflow.com/questions/13424096/contact-form-is-not-sending-email-to-my-gmail-acount
            SmtpServer.EnableSsl = True
            SmtpServer.DeliveryMethod = SmtpDeliveryMethod.Network
            SmtpServer.UseDefaultCredentials = False
            SmtpServer.Credentials = New Net.NetworkCredential(strUsername, strPassword)
            ' Citation end.
            mail = New MailMessage()
            mail.From = New MailAddress(strFrom)
            mail.To.Add(strTO)
            mail.Subject = strSubject
            mail.Body = strBody
            ' Add Attachment
            ' Code from: http://vb.net-informations.com/communications/vb-email-attachment.htm
            ' Reference for excel: https://www.tutorialspoint.com/vb.net/vb.net_excel_sheet.htm
            ' Reference for PDF And Easy Report RDLC: https://www.youtube.com/watch?v=HX8hG29s3r8
            Dim attachment As System.Net.Mail.Attachment
            attachment = New System.Net.Mail.Attachment(strAttachmentPath)
            mail.Attachments.Add(attachment)

            SmtpServer.Send(mail)
            MsgBox("mail send")
            Return 0
        Catch ex As Exception
            MsgBox(ex.ToString)
            Return ex.Message.Length
        End Try
    End Function

End Module
