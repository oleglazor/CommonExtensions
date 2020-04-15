Imports System.Net.Mail

Public Interface ISmtpClientWrapper 
    Inherits IDisposable

    Sub Send(message As MailMessage)
    Property Host As String
    Property Port As Integer

End Interface
Friend Class SmtpClientWrapper
    Implements ISmtpClientWrapper
    
    Private _isDisposed As Boolean
    Private ReadOnly _smtpClient As Net.Mail.SmtpClient
    
    Public Property Host As String Implements ISmtpClientWrapper.Host 
    Public Property Port As Integer Implements ISmtpClientWrapper.Port
    
    Public Sub New(smtpClient As Net.Mail.SmtpClient)
        _smtpClient = smtpClient
        Host = smtpClient.Host
        Port = smtpClient.Port
    End Sub

    Public Sub Send(message As MailMessage) Implements ISmtpClientWrapper.Send
        _smtpClient.Send(message)
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
