Imports System.Net.Mail

Public Class MailHelper
    ''' <summary>
    ''' Sends an mail message
    ''' </summary>
    ''' <param name="smtpserver">SMTP Server</param>
    ''' <param name="from">Sender address</param>
    ''' <param name="recipients">Recepient address</param>
    ''' <param name="bcc">Bcc recepient</param>
    ''' <param name="cc">Cc recepient</param>
    ''' <param name="subject">Subject of mail message</param>
    ''' <param name="body">Body of mail message</param>
    Public Shared Sub SendMailMessage(ByVal smtpserver As String, ByVal from As String, ByVal recipients As String, ByVal bcc As String, ByVal cc As String, _
        ByVal subject As String, ByVal body As String)
        ' Instantiate a new instance of MailMessage
        Dim mMailMessage As New MailMessage()

        ' Set the sender address of the mail message
        mMailMessage.From = New MailAddress(from)
        ' Set the recepient address of the mail message
        ' Updated below for multiple recipients
        'mMailMessage.To.Add(New MailAddress(recepient))

        ' Set the recepient address of the mail message
        ' Changed to accomodate multiple recipients
        Dim arrayRecipients() As String = recipients.Split(",")
        Dim rcpnts As String

        For Each rcpnts In arrayRecipients
            mMailMessage.To.Add(New MailAddress(rcpnts))
        Next

        ' Check if the bcc value is null or an empty string
        If Not bcc Is Nothing And bcc <> String.Empty Then
            ' Set the Bcc address of the mail message
            mMailMessage.Bcc.Add(New MailAddress(bcc))
        End If

        ' Check if the cc value is null or an empty value
        If Not cc Is Nothing And cc <> String.Empty Then
            ' Set the CC address of the mail message
            mMailMessage.CC.Add(New MailAddress(cc))
        End If

        ' Set the subject of the mail message
        mMailMessage.Subject = subject
        ' Set the body of the mail message
        mMailMessage.Body = body

        ' Secify the format of the body as HTML
        mMailMessage.IsBodyHtml = True
        ' Set the priority of the mail message to normal
        mMailMessage.Priority = MailPriority.Normal

        ' Instantiate a new instance of SmtpClient
        Dim mSmtpClient As New SmtpClient(smtpserver)
        ' Send the mail message
        mSmtpClient.Send(mMailMessage)
    End Sub
End Class
