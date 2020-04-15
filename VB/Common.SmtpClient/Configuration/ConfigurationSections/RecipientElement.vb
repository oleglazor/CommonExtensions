Imports System.Configuration

Public Class Recipient
    Inherits ConfigurationElement

    Private Const EmailAttributeName = "email"
    Private Const DisplayNameAttributeName = "displayName"
    Private Const GroupsAttributeName = "groups"

    Public Sub New()
    End Sub

    Public Sub New(ByVal email As String, ByVal displayName As String, groups As string)
        Me.Email = email
        Me.DisplayName = displayName
        Me.Groups = groups
    End Sub

    <ConfigurationProperty(EmailAttributeName, IsRequired:=True, IsKey:=True)>
    Public Property Email As String
        Get
            Return Me(EmailAttributeName)
        End Get
        Set(ByVal value As String)
            Me(EmailAttributeName) = value
        End Set
    End Property

    <ConfigurationProperty(GroupsAttributeName, IsRequired:=True, IsKey:=False)>
    Public Property Groups As String
        Get
            Return CStr(Me(GroupsAttributeName))
        End Get
        Set(ByVal value As String)
            Me(GroupsAttributeName) = value
        End Set
    End Property

    <ConfigurationProperty(DisplayNameAttributeName, IsRequired:=False, IsKey:=False)>
    Public Property DisplayName As String
        Get
            Return CStr(Me(DisplayNameAttributeName))
        End Get
        Set(ByVal value As String)
            Me(DisplayNameAttributeName) = value
        End Set
    End Property

End Class