Imports System.Configuration
Imports System.Net.Mail
Imports Common.Extensions

Public Class SmtpSettingsSection
    Inherits ConfigurationSection

    Private Const ServersSectionName = "SmtpServers"
    Private Const SendersSectionName = "Senders"
    Private Const RecipientGroupsName = "Recipients"
    Public Sub New()
        'need for unit tests
    End Sub

    Public Sub New(smtpServers As SmtpServersCollection)
        Me.SmtpServers.Clear()

        For Each smtpServer As SmtpServer In smtpServers
            Me.SmtpServers.Add(smtpServer)
        Next
    End Sub

    <ConfigurationProperty(ServersSectionName, IsDefaultCollection:=False)>
    Public ReadOnly Overloads Property SmtpServers As SmtpServersCollection
        Get
            Return CType(Item(ServersSectionName), SmtpServersCollection)
        End Get
    End Property

    <ConfigurationProperty(SendersSectionName, IsDefaultCollection:=False)>
    Public ReadOnly Overloads Property Senders As SendersCollection
        Get
            Return CType(Item(SendersSectionName), SendersCollection)
        End Get
    End Property

    <ConfigurationProperty(RecipientGroupsName, IsDefaultCollection:=False)>
    Public ReadOnly Overloads Property Recipients As RecipientsCollection
        Get
            Return CType(Item(RecipientGroupsName), RecipientsCollection)
        End Get
    End Property

    Public Function GetByRecipientsGroup(groups As IEnumerable(Of RecipientGroup)) As IEnumerable(Of MailAddress)
        Return GetByRecipientsGroup(groups.Select(Function(g) g.ToString()).ToArray())
    End Function

    Public Function GetByRecipientsGroup(ParamArray groupNames As String()) As IEnumerable(Of MailAddress)
        Dim recipientsSection = TryCast(ConfigurationManager.GetSection(Me.GetType().Name), SmtpSettingsSection)
        Dim recipients = New Dictionary(Of String,String)

        For Each recipient As Recipient In recipientsSection.Recipients
            recipient.Groups.Split(",").ToList().ForEach(
                Function(group)
                    If groupNames.Contains(group) And Not recipients.ContainsKey(recipient.Email) Then
                        recipients.Add(recipient.Email, recipient.DisplayName)
                    End If
                End Function
            )
        Next

        Return recipients.Select(Function(r) new MailAddress(r.Key, r.Value))
    End Function

    Public Function GetSenderByName(sender As Senders) As MailAddress
        Return GetSenderByName(sender.ToString())
    End Function

    Public Function GetSenderByName(name As String) As MailAddress
        Dim sendersSection = TryCast(ConfigurationManager.GetSection(Me.GetType().Name), SmtpSettingsSection)

        For Each sender As Sender In sendersSection.Senders
            If sender.Name.Equals(name) Then
                Return new MailAddress(sender.Email, sender.DisplayName)
            End If
        Next

        Throw new ArgumentException($"Can't find sender with name {name}")
    End Function
End Class