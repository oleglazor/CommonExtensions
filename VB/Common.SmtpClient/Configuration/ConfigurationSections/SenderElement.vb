Imports System.Configuration

Public Class Sender
    Inherits ConfigurationElement

    Private Const EmailAttributeName = "email"
    Private Const DisplayNameAttributeName = "displayName"
    Private Const NameAttributeName = "name"

    Public Sub New()
    End Sub

    Public Sub New(name As String, email As String, displayName As String)
        Me.Name = name
        Me.Email = email
        Me.DisplayName = displayName
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

    <ConfigurationProperty(DisplayNameAttributeName, IsRequired:=False, IsKey:=False)>
    Public Property DisplayName As String
        Get
            Return CStr(Me(DisplayNameAttributeName))
        End Get
        Set(ByVal value As String)
            Me(DisplayNameAttributeName) = value
        End Set
    End Property

    <ConfigurationProperty(NameAttributeName, IsRequired:=True, IsKey:=False)>
    Public Property Name As String
        Get
            Return CStr(Me(NameAttributeName))
        End Get
        Set(ByVal value As String)
            Me(NameAttributeName) = value
        End Set
    End Property
End Class