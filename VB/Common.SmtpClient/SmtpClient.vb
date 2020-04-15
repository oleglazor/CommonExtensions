Imports System.Net.Mail
Imports Common.Extensions

Public Interface ISmtpClient 
    Inherits IDisposable
    
    Sub Send(emailMessage As MailMessage) 

    Sub Send(
         toEmails As IEnumerable(Of String), 
         mailBody As String, subject As String, 
         ByVal Optional fromEmail As String = Nothing, 
         ByVal Optional replyTo As IEnumerable(Of String) = Nothing, 
         ByVal Optional isBodyHtml As Boolean = False,
         ByVal Optional attachments As IEnumerable(Of Attachment) = Nothing) 

    Sub Send(
             toEmails As IEnumerable(Of String), 
             mailBody As String, 
             subject As String, 
             ByVal Optional fromEmail As Senders = Nothing, 
             ByVal Optional replyTo As IEnumerable(Of RecipientGroup) = Nothing, 
             ByVal Optional isBodyHtml As Boolean = False,
             ByVal Optional attachments As IEnumerable(Of Attachment) = Nothing)

    Sub Send(
        recipients As RecipientGroup, 
        mailBody As String, 
        subject As String, 
        ByVal Optional fromEmail As Senders = Senders.DoNotReply, 
        ByVal Optional replyTo As RecipientGroup = Nothing, 
        ByVal Optional isBodyHtml As Boolean = False,
        ByVal Optional attachments As IEnumerable(Of Attachment) = Nothing)

    Sub Send(
             toEmails As IEnumerable(Of MailAddress), 
             mailBody As String, 
             subject As String, 
             ByVal Optional fromEmail As MailAddress = Nothing, 
             ByVal Optional replyTo As IEnumerable(Of MailAddress) = Nothing, 
             ByVal Optional isBodyHtml As Boolean = False,
             ByVal Optional attachments As IEnumerable(Of Attachment) = Nothing)

    Sub Send(
             toEmails As RecipientGroup, 
             mailBody As String, 
             subject As String, 
             ByVal Optional fromEmail As Senders = Nothing, 
             ByVal Optional replyTo As IEnumerable(Of RecipientGroup) = Nothing, 
             ByVal Optional isBodyHtml As Boolean = False,
             ByVal Optional attachments As IEnumerable(Of Attachment) = Nothing)

End Interface

''' <summary>
''' this is featured wrapper around Net.Mail.SmtpClient. It attempt to send email via available servers one by one and will stop after any successful attempt.
''' it expect section in configuration file
''' example:
'''    <configSections>
'''        <section name="SmtpSettingsSection" type="NHSScotland.Payroll.Utilities.Extensions.SmtpSettingsSection, NHSScotland.Payroll.Utilities.Extensions"/>
'''    </configSections>
'''    <SmtpSettingsSection>
'''        <SmtpServers>
'''            <add host="address.A" port="25" switchOn="true"></add>
'''            ...........................................................
'''            <add host="address.B" port="25" switchOn="true"></add>
'''            <add host="address.C" port="25" switchOn="false"></add>
'''    </SmtpServers>
'''  <Senders>
'''     <add name="ShortName" email="email@domain.d" displayName="DISPLAY NAME"/>
''' ....................................
'''  </Senders>
'''  <Recipients>
'''     <add email="email@domain.d" displayName="display name" groups="Group1,Group2,...,GroupN"/>
'''     .......................................
'''  </Recipients>
''' </SmtpSettingsSection>
'''     ..................
''' </summary>
Public Class SmtpClient
    Implements ISmtpClient

    Private _isDisposed As Boolean
    Private ReadOnly _smtpServersConfigurationSection As SmtpSettingsSection
    Private ReadOnly _smtpServers As SmtpServersCollection
    Private ReadOnly _areAllOff As Boolean
    Private ReadOnly _configurationProvider As IConfigurationProvider
    Private ReadOnly _smtpClient As ISmtpClientWrapper

    ''' <summary>
    ''' this constructor will try to get SmtpSettingsSection from web config via IConfigurationProvider.Get
    ''' </summary>
    ''' <param name="configurationProvider">require implementation of IConfiguration provider</param>
    ''' <param name="smtpServerForTesting">smtpServerForTesting needed just for testing. Do not pass it from any production code</param>
    Public Sub New(configurationProvider As IConfigurationProvider, ByRef Optional smtpServerForTesting As ISmtpClientWrapper = Nothing)
        _configurationProvider = configurationProvider
        _smtpClient = If(smtpServerForTesting IsNot Nothing, smtpServerForTesting, New SmtpClientWrapper(New Net.Mail.SmtpClient()))
        _smtpServersConfigurationSection = _configurationProvider.Get(Of SmtpSettingsSection)(GetType(SmtpSettingsSection).Name, SettingType.ConfigSection)
        Check.ThrowIfNullOrEmpty(_smtpServersConfigurationSection?.SmtpServers, messageOverride := "No smtp servers configured")
        _smtpServers = _smtpServersConfigurationSection.SmtpServers
        '_senders = _smtpServersConfigurationSection.Senders
        '_recipients = _smtpServersConfigurationSection.Recipients
        _areAllOff = Not (from server As SmtpServer in _smtpServers where server.SwitchOn select server).Any()
    End Sub

    Public Sub Send(emailMessage As MailMessage) Implements ISmtpClient.Send
        Check.ThrowIfNullOrEmpty(emailMessage)
        Check.ThrowIfNullOrEmpty(emailMessage.From)
        Check.ThrowIfNullOrEmpty(emailMessage.To)
        Check.ThrowIfNullOrEmpty(emailMessage.Body)
        Check.ThrowIfNullOrEmpty(emailMessage.Subject)

        Send({emailMessage.To}.AsEnumerable(), emailMessage.Body, emailMessage.Subject, emailMessage.From, emailMessage.ReplyToList, emailMessage.IsBodyHtml)
    End Sub

    Public Sub Send(
                    recipients As RecipientGroup, 
                    mailBody As String, 
                    subject As String, 
                    ByVal Optional fromEmail As Senders = Senders.DoNotReply, 
                    ByVal Optional replyTo As RecipientGroup = Nothing, 
                    ByVal Optional isBodyHtml As Boolean = False,
                    ByVal Optional attachments As IEnumerable(Of Attachment) = Nothing) Implements ISmtpClient.Send

        Send(recipients, mailBody, subject, fromEmail, replyTo, isBodyHtml, attachments)
    End Sub

    Public Sub Send(
                    toEmails As IEnumerable(Of String), 
                    mailBody As String, 
                    subject As String, 
                    ByVal Optional fromEmail As Senders = Nothing, 
                    ByVal Optional replyTo As IEnumerable(Of RecipientGroup) = Nothing, 
                    ByVal Optional isBodyHtml As Boolean = False,
                    ByVal Optional attachments As IEnumerable(Of Attachment) = Nothing) Implements ISmtpClient.Send

        Dim toEmailsList = New List(Of MailAddress)
        toEmails.ToList().ForEach(Sub(a) toEmailsList.Add(New MailAddress(a)))
        
        Dim sender = _smtpServersConfigurationSection.GetSenderByName(fromEmail)
        Check.ThrowIfNullOrEmpty(sender)

        Dim replyToList As IEnumerable(Of MailAddress) = Nothing
        If Not Check.IsNullOrEmpty(replyTo) Then
            replyToList = _smtpServersConfigurationSection.GetByRecipientsGroup(replyTo)
        End If
        
        Send(toEmailsList, mailBody, subject, sender, replyToList, isBodyHtml, attachments)

    End Sub

    Public Sub Send(
                    toEmails As RecipientGroup, 
                    mailBody As String, 
                    subject As String, 
                    ByVal Optional fromEmail As Senders = Nothing, 
                    ByVal Optional replyTo As IEnumerable(Of RecipientGroup) = Nothing, 
                    ByVal Optional isBodyHtml As Boolean = False,
                    ByVal Optional attachments As IEnumerable(Of Attachment) = Nothing) Implements ISmtpClient.Send

        Dim toList = _smtpServersConfigurationSection.GetByRecipientsGroup(toEmails)
        Check.ThrowIfNullOrEmpty(replyTo)

        Dim sender = _smtpServersConfigurationSection.GetSenderByName(fromEmail)
        Check.ThrowIfNullOrEmpty(sender)

        Dim replyToList = Nothing
        If Not Check.IsNullOrEmpty(replyTo) Then
            replyToList = _smtpServersConfigurationSection.GetByRecipientsGroup(replyto)
        End If

        Send(toList, mailBody, subject, sender, replyToList, isBodyHtml, attachments)

    End Sub

    Public Sub Send(
                    toEmails As IEnumerable(Of String), 
                    mailBody As String, 
                    subject As String, 
                    ByVal Optional fromEmail As String = Nothing, 
                    ByVal Optional replyTo As IEnumerable(Of String) = Nothing, 
                    ByVal Optional isBodyHtml As Boolean = False,
                    ByVal Optional attachments As IEnumerable(Of Attachment) = Nothing) Implements ISmtpClient.Send

        Check.ThrowIfNullOrEmpty(fromEmail, messageOverride := "sender email can't be empty")

        Dim toEmailsList = New List(Of MailAddress)
        toEmails.ToList().ForEach(Sub(a) toEmailsList.Add(New MailAddress(a)))
        
        
        Dim replyToList = Nothing
        If Not Check.IsNullOrEmpty(replyTo) Then
            replyToList = New List(Of MailAddress)
            replyTo.ToList().ForEach(Sub(r) replyToList.Add(New MailAddress(r)))
        End If

        Send(toEmailsList,mailBody,subject,new MailAddress(fromEmail), replyToList, isBodyHtml, attachments)
    End Sub

    Public Sub Send(
                    toEmails As IEnumerable(Of MailAddress), 
                    mailBody As String, 
                    subject As String, 
                    ByVal Optional fromEmail As MailAddress = Nothing, 
                    ByVal Optional replyTo As IEnumerable(Of MailAddress) = Nothing, 
                    ByVal Optional isBodyHtml As Boolean = False,
                    ByVal Optional attachments As IEnumerable(Of Attachment) = Nothing) Implements ISmtpClient.Send
        
        If _areAllOff Then Return

        Check.ThrowIfNullOrEmpty(subject, messageOverride := "Please provide a subject")
        Check.ThrowIfNullOrEmpty(fromEmail, messageOverride := "Message sender should be provided")
        Check.ThrowIfNullOrEmpty(toEmails, messageOverride := "Message recipient(s) should be provided")
        Check.ThrowIfNullOrEmpty(mailBody, messageOverride := "Message would be provided")

        Dim emailMsg As New MailMessage()
        emailMsg.IsBodyHtml = isBodyHtml
        emailMsg.Subject = subject
        emailMsg.From = fromEmail
        emailMsg.Body = mailBody

        For Each toEmail As MailAddress In toEmails
            emailMsg.To.Add(toEmail)
        Next

        If Not Check.IsNullOrEmpty(replyTo) Then
            For Each replyToEmail As MailAddress In replyTo
                emailMsg.ReplyToList.Add(replyToEmail)
            Next
        End If

        If Not Check.IsNullOrEmpty(attachments) Then
            For Each attachment As Attachment In attachments
                emailMsg.Attachments.Add(attachment)
            Next
        End If
        
        SendViaFirstAvailableSmtpServer(emailMsg)

    End Sub

    Private Sub SendViaFirstAvailableSmtpServer(mailMessage As MailMessage)
        Dim isSent As Boolean
        Dim exceptions = New List(Of Exception)()
        For Each smtpServer As SmtpServer In _smtpServers
            Try
                If Not smtpServer.SwitchOn
                    Continue For
                End If

                _smtpClient.Host = smtpServer.Host
                _smtpClient.Port = smtpServer.Port
                
                For i = 0 To 3
                    Try
                        _smtpClient.Send(mailMessage)
                        isSent = True

                        Exit For
                    Catch ex As Exception
                        exceptions.Add(ex)
                        'log exception
                    End Try
                Next
                
                Exit For
            Catch ex As Exception
                exceptions.Add(ex)
                'log exception
            End Try
        Next

        mailMessage.Dispose()

        If Not isSent Then
            If Not Check.IsNullOrEmpty(exceptions) Then
                Throw new AggregateException(exceptions)
            Else 
                Throw New Exception("Application error: can't sent email")
            End If
        End If

    End Sub

    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not _isDisposed Then
            _smtpClient?.Dispose()
        End If

        _isDisposed = True
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub

    Protected Overrides Sub Finalize()
        Dispose(False)
    End Sub

End Class
