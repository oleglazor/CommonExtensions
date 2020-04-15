Imports System.Configuration

Public Class SmtpServer
    Inherits ConfigurationElement

    Private Const HostAttributeName = "host"
    Private Const PortAttributeName = "port"
    Private Const SwitchOnAttributeName = "switchOn"

    Public Sub New()
    End Sub

    Public Sub New(ByVal host As String, ByVal port As Integer, switchOn As Boolean)
        Me.Host = host
        Me.Port = port
        Me.SwitchOn = switchOn
    End Sub

    <ConfigurationProperty(HostAttributeName, IsRequired:=True, IsKey:=True)>
    Public Property Host As String
        Get
            Return Me(HostAttributeName)
        End Get
        Set
            Me(HostAttributeName) = value
        End Set
    End Property

    <ConfigurationProperty(PortAttributeName, IsRequired:=True, IsKey:=False)>
    Public Property Port As Integer
        Get
            Return CStr(Me(PortAttributeName))
        End Get
        Set
            Me(PortAttributeName) = value
        End Set
    End Property

    <ConfigurationProperty(SwitchOnAttributeName, IsRequired:=True, IsKey:=False)>
    Public Property SwitchOn As Boolean
        Get
            Return CStr(Me(SwitchOnAttributeName))
        End Get
        Set
            Me(SwitchOnAttributeName) = value
        End Set
    End Property
End Class